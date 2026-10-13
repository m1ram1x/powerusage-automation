"""Plain-text Telegram alerts; credentials never belong in source files."""

from __future__ import annotations

import json
import os
import re
import ssl
import time
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import HTTPRedirectHandler, HTTPSHandler, Request, build_opener


class TelegramError(ValueError):
    """Safe to display: never includes a token, request URL or API response."""


@dataclass(frozen=True)
class Alert:
    source: str
    filename: str
    row: int
    description: str


def messages(alerts: list[Alert]) -> list[str]:
    """Group by actual PDU identity and split long messages before API limits."""
    grouped = defaultdict(list)
    for alert in dict.fromkeys(alerts):
        grouped[alert.source].append(alert)
    result = []
    for source, items in sorted(grouped.items()):
        # Bound untrusted file labels; plain text means no HTML/Markdown parsing.
        heading = source.replace("\n", " ").replace("\r", " ")[:200] + ":\n"
        text = heading
        for item in items:
            filename = item.filename.replace("\n", " ").replace("\r", " ")[:200]
            line = f"\n{item.description}\nФайл: {filename}\nСтрока: {item.row}\n\n"
            if text != heading and len(text) + len(line) > 1800:
                result.append(text.rstrip())
                text = heading
            # 1800 Unicode code points also fit within 4096 UTF-16 units.
            while line:
                room = 1800 - len(text)
                text += line[:room]
                line = line[room:]
                if line or len(text) == 1800:
                    result.append(text.rstrip())
                    text = heading
        if text != heading:
            result.append(text.rstrip())
    return result


def credentials(env_file: Path | None = None, *, require_chat: bool = True) -> tuple[str, str]:
    """Read two literal values; never source/evaluate an environment file."""
    values = {}
    allowed = {"TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID"}
    if env_file is not None:
        try:
            if os.name != "nt" and env_file.stat().st_mode & 0o077:
                raise TelegramError("Protect the credential file first: chmod 600 <file>")
            contents = env_file.read_text(encoding="utf-8-sig")
        except OSError:
            raise TelegramError("Cannot read the Telegram configuration file") from None
        for line in contents.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            key, separator, value = line.partition("=")
            if not separator or key.strip() not in allowed or key.strip() in values:
                raise TelegramError("Telegram configuration: expected two unique KEY=value settings")
            values[key.strip()] = value.strip()
    # A chosen file is authoritative, to avoid silently selecting another chat.
    if env_file is None:
        values = {key: os.environ.get(key, "").strip() for key in allowed}
    token, chat = values.get("TELEGRAM_BOT_TOKEN", ""), values.get("TELEGRAM_CHAT_ID", "")
    if not re.fullmatch(r"[0-9]{5,20}:[A-Za-z0-9_-]{20,128}", token):
        raise TelegramError("Set a valid TELEGRAM_BOT_TOKEN in your local configuration")
    if (require_chat or chat) and not re.fullmatch(r"-?[0-9]{1,20}|@[A-Za-z][A-Za-z0-9_]{4,31}", chat):
        raise TelegramError("Set TELEGRAM_CHAT_ID to your numeric chat ID or channel @username")
    return token, chat


class NoRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # The URL includes a credential: do not follow redirects to another host.
        return None


def api_call(token: str, method: str, payload: dict) -> object:
    if method not in {"sendMessage", "getUpdates"}:
        raise TelegramError("Unsupported Telegram API operation")
    opener = build_opener(NoRedirects(), HTTPSHandler(context=ssl.create_default_context()))
    request = Request(f"https://api.telegram.org/bot{token}/{method}",
                      data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                      headers={"Content-Type": "application/json"}, method="POST")
    try:
        with opener.open(request, timeout=15) as response:
            raw = response.read(1_000_001)
        if len(raw) > 1_000_000:
            raise TelegramError("Telegram returned an oversized response")
        body = json.loads(raw)
        if not isinstance(body, dict) or body.get("ok") is not True:
            raise TelegramError("Telegram rejected the request; check token and chat permissions")
        return body.get("result")
    except HTTPError as exc:
        # Do not print exc: its URL contains the token. Don't repeat ambiguous sends.
        code = exc.code
        exc.close()
        raise TelegramError(f"Telegram HTTP {code}; check connection, credentials and chat permissions") from None
    except (URLError, OSError, ValueError) as exc:
        if isinstance(exc, TelegramError):
            raise
        raise TelegramError("Telegram request failed; check HTTPS connectivity and retry manually") from None


def send_alerts(alerts: list[Alert], token: str, chat: str) -> int:
    sent = 0
    for text in messages(alerts):
        if sent:
            # Private chat: 1/sec. Groups: no more than 20/minute.
            time.sleep(3.1 if chat.startswith(("-", "@")) else 1.1)
        try:
            api_call(token, "sendMessage", {"chat_id": chat, "text": text})
        except TelegramError as exc:
            raise TelegramError(f"{exc}. Confirmed sent: {sent}; remaining delivery is not confirmed") from None
        sent += 1
    return sent

def send_excel(path: Path, token: str, chat: str) -> None:
    from uuid import uuid4

    try:
        with path.open("rb") as stream:
            data = stream.read(50_000_001)
    except OSError:
        raise TelegramError("Не удалось прочитать итоговый Excel") from None

    if len(data) > 50_000_000:
        raise TelegramError("Excel слишком большой для отправки")

    boundary = uuid4().hex
    filename = re.sub(r"[^A-Za-z0-9_.-]", "_", path.name)

    body = (
        f"--{boundary}\r\n"
        'Content-Disposition: form-data; name="chat_id"\r\n\r\n'
        f"{chat}\r\n"
        f"--{boundary}\r\n"
        'Content-Disposition: form-data; name="document"; '
        f'filename="{filename}"\r\n'
        "Content-Type: application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet\r\n\r\n"
    ).encode("utf-8")

    body += data + f"\r\n--{boundary}--\r\n".encode("ascii")

    request = Request(
        f"https://api.telegram.org/bot{token}/sendDocument",
        data=body,
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}"
        },
        method="POST",
    )

    opener = build_opener(
        NoRedirects(),
        HTTPSHandler(context=ssl.create_default_context()),
    )

    # Пауза после возможного отправленного предупреждения.
    time.sleep(3.1 if chat.startswith(("-", "@")) else 1.1)

    try:
        with opener.open(request, timeout=60) as response:
            raw = response.read(1_000_001)

        if len(raw) > 1_000_000:
            raise TelegramError("Слишком большой ответ Telegram")

        result = json.loads(raw)
        if not isinstance(result, dict) or result.get("ok") is not True:
            raise TelegramError("Telegram не подтвердил отправку Excel")

    except HTTPError as exc:
        code = exc.code
        exc.close()
        raise TelegramError(f"Ошибка отправки Excel: HTTP {code}") from None
    except (URLError, OSError, ValueError) as exc:
        if isinstance(exc, TelegramError):
            raise
        raise TelegramError(
            "Отправка Excel не подтверждена; проверь соединение"
        ) from None