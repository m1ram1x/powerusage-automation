# Troubleshooting

## Docker is missing or stopped

Check `docker version` and `docker info`. On Ubuntu, install Docker through the official guide linked in the README or your distribution's package manager, then start its service. If `apt install docker.io` reports 404 downloads, run `sudo apt update` successfully before retrying the installation.

On Windows, start Docker Desktop and use Linux-container mode. Verify:

```powershell
docker info --format '{{.OSType}}'
```

If Docker reports missing WSL/Virtual Machine Platform components, follow [Docker's WSL guide](https://docs.docker.com/desktop/features/wsl/) and [Microsoft's WSL installation instructions](https://learn.microsoft.com/en-us/windows/wsl/install). Windows-feature changes require an administrator window and may require a reboot. A Windows servicing error is a system issue, not a CSV parser error; record its exact code rather than modifying the registry blindly.

## No CSV files found

The repository intentionally ships an empty `input`. Add your CSV exports there, or use `bash run.sh --demo` / the Windows launcher with `-Demo`. Demo files do not need to be copied into your private input folder.

## Template or mount path does not exist

Keep `templates/2026_PowerUsageTable_template.xlsx` in the extracted repository. Run launchers from the download; they resolve paths relative to their own location. The original file name is case-sensitive on Linux.

## No usable temporary directory / permission denied

The launchers provide a writable `/tmp` tmpfs and a writable host `output` while keeping the container root read-only. If writing your own command, include:

```text
--tmpfs /tmp:rw,noexec,nosuid,nodev,size=128m,mode=1777
```

For very large batches, increase the tmpfs size after assessing available RAM. On Linux, run the container using your UID/GID and ensure you own the `output` directory. Avoid making an entire project world-writable. On Windows, verify Docker Desktop can access the selected host folders.

## Excel shows older data

Find the `Saved ...` line and open that exact local file from `output`. `/app/output` inside Docker refers to the mounted host `output`. Close any old spreadsheet window and check whether your viewer uploaded a separate cloud copy. Use local Python for a read-only check:

```bash
python -c "from openpyxl import load_workbook; w=load_workbook('output/YOUR_REPORT.xlsx',read_only=True); print(*w['YaIDC'].iter_rows(min_row=1,max_row=5,max_col=7,values_only=True),sep='\n'); w.close()"
```

Replace `YOUR_REPORT.xlsx` with the saved filename. Editing an output is not an input change: a successful rerun regenerates its values from CSV and the template. A failed run can leave the previous report on disk.

## Only one output row

One row represents one rack/date pair, not one CSV. Two line-A/line-B reports for the same rack/date combine into one row. A file containing several rack blocks can contribute several rows.

## Telegram says the token is missing

Create a local `telegram.env` using the empty example, save the token and chat ID without quotes, and enable `--telegram` / `-Telegram`. For a Python/PyCharm run, pass `--telegram --telegram-env telegram.env` as script parameters and set the working directory to the project root.

`telegram_setup.py` also requires `--telegram-env telegram.env` when using that file. It does not automatically read a dotenv file when no path is supplied.

## Telegram warnings: 0

Zero warnings can be correct. The document upload is separate and occurs after successful processing; look for the Excel delivery line. A quiet warning run alone is not proof of a successful API connection.

## Telegram request failed

Check the bot's `/start`, chat ID, chat permissions and HTTPS connectivity. HTTP 401 usually concerns the token; HTTP 403 can concern blocked/forbidden chat access; HTTP 429 means rate limiting. Do not disable certificate checks or publish the token/complete request URL. The local report is retained when sending fails. Confirm actual chat delivery before retrying after a timeout.

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Processing and requested delivery completed |
| 1 | Input, template, output or launcher error |
| 2 | Telegram configuration/delivery error, without a prior data error; also CLI usage errors |

When data processing and Telegram delivery both fail, the application keeps code 1 and displays the delivery failure separately. No report is sent when the application has recorded a processing error.
