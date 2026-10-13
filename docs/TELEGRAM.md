# Telegram setup

## Create your bot

Open [@BotFather](https://t.me/BotFather), send `/newbot`, and follow its prompts. Save the token locally; treat it as a password. Then open **your new bot** and send `/start`.

[Telegram's official bot tutorial](https://core.telegram.org/bots/tutorial#obtain-your-bot-token).

## Create settings

Copy `telegram.env.example` to `telegram.env` without overwriting an existing settings file, as shown in the main README. Enter the token after `TELEGRAM_BOT_TOKEN=`. Leave `TELEGRAM_CHAT_ID=` blank during chat-ID discovery. Use literal values without quotes; the file is never executed as shell code.

On Linux, run `chmod 600 telegram.env`. On Windows, keep the file in your own user directory and restrict access through its **Properties → Security** settings. The file is plain text, not an encrypted secret store.

## Obtain the chat ID with Docker

First build the image using the demo launcher, or:

```bash
docker build -t powerusage-automation:local .
```

On Linux, from the project folder, paste this entire block:

```bash
sudo docker run --rm \
  --user "$(id -u):$(id -g)" \
  --read-only --cap-drop ALL --security-opt no-new-privileges \
  --tmpfs /tmp:rw,noexec,nosuid,nodev,size=128m,mode=1777 \
  --mount "type=bind,source=$(pwd)/telegram.env,target=/run/secrets/telegram.env,readonly" \
  powerusage-automation:local \
  --setup-chat --telegram-env /run/secrets/telegram.env
```

On Windows PowerShell:

```powershell
$projectPath = (Get-Location).Path
docker run --rm --user 10001:10001 --read-only --cap-drop ALL --security-opt no-new-privileges --tmpfs /tmp:rw,noexec,nosuid,nodev,size=128m,mode=1777 --mount "type=bind,source=$projectPath/telegram.env,target=/run/secrets/telegram.env,readonly" powerusage-automation:local --setup-chat --telegram-env /run/secrets/telegram.env
```

This reads recent updates from your bot, without sending anything or printing the token. Expected output:

```text
Chat ID: 123456789 (private)
```

Put your actual number after `TELEGRAM_CHAT_ID=` and save the file. If the output says `No recent chats`, send `/start` to your own bot again, then retry. An already configured webhook can prevent `getUpdates`; use a dedicated bot or review its webhook setup. The helper does not remove a webhook automatically.

Alternatively, with local Python:

```bash
.venv/bin/python telegram_setup.py --telegram-env telegram.env
```

Windows equivalent:

```powershell
.\.venv\Scripts\python.exe telegram_setup.py --telegram-env telegram.env
```

[Telegram getUpdates documentation](https://core.telegram.org/bots/api#getupdates).

## Send reports

Put CSV files in `input`, then run `bash run.sh --telegram` on Linux or `run.ps1 -Telegram` using the PowerShell command in the README.

The generated Excel is sent after successful processing, including when warning count is zero. CSV files are not sent. The report and warnings disclose their contained measurements and rack information to the configured chat, so choose a chat you are authorized to use.

Document uploads are limited to 50,000,000 bytes by this application. See [sendDocument](https://core.telegram.org/bots/api#senddocument). If a request times out, the program does not retry automatically; check the chat before repeating the run to avoid duplicate delivery.

Never publish `telegram.env`, paste a real token into an issue, screenshot, source file or terminal command. If a token is exposed, revoke it through BotFather and replace the local value. CI does not need Telegram secrets.
