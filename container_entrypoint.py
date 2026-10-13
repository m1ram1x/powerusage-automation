"""Stage a mounted secret privately; Windows bind mounts do not preserve chmod.

The host's ACL/file permissions protect the original. The temporary copy is
owner-only and is removed after processing. Never print its contents.
"""

import os
import sys
import tempfile
from pathlib import Path

from main import main as process_batch
from telegram_setup import main as find_chat_ids


def main(argv=None):
    arguments = list(sys.argv[1:] if argv is None else argv)
    setup_chat = "--setup-chat" in arguments
    if setup_chat:
        arguments.remove("--setup-chat")
    private_path = None
    try:
        if "--telegram" in arguments or setup_chat:
            for index, argument in enumerate(arguments):
                if argument == "--telegram-env" and index + 1 < len(arguments):
                    source = Path(arguments[index + 1])
                    value_index = index + 1
                    inline = False
                elif argument.startswith("--telegram-env="):
                    source = Path(argument.split("=", 1)[1])
                    value_index = index
                    inline = True
                else:
                    continue
                with source.open("rb") as stream:
                    data = stream.read(8193)
                if len(data) > 8192:
                    print("TELEGRAM: credential file is too large", file=sys.stderr)
                    return 2
                with tempfile.NamedTemporaryFile(prefix="powerusage-secret-", suffix=".env",
                                                  delete=False) as secret:
                    private_path = Path(secret.name)
                    os.chmod(private_path, 0o600)
                    secret.write(data)
                del data
                arguments[value_index] = ("--telegram-env=" if inline else "") + str(private_path)
                break
        return find_chat_ids(arguments) if setup_chat else process_batch(arguments)
    except OSError:
        print("ERROR: cannot prepare the mounted Telegram settings; check mounts and /tmp", file=sys.stderr)
        return 2
    finally:
        if private_path is not None:
            private_path.unlink(missing_ok=True)


if __name__ == "__main__":
    raise SystemExit(main())
