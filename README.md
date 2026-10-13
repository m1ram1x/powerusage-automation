# Power Usage Automation

Combine PDU CSV exports into a single Excel report and optionally deliver the report and alerts through a Telegram bot.

Each Excel row represents **one rack on one reporting date**. Readings from the two power lines are matched by rack and date, rather than by filename or row order.

```text
PDU CSV files + Excel template
              ↓
      Python application in Docker
              ↓
   Excel report + optional Telegram delivery
```

## Features

- Read all `.csv` files directly inside the selected input folder.
- Identify the report source from CSV metadata, with a filename fallback.
- Merge Average and Max readings into PDU1/PDU2 columns.
- Create a fresh report while preserving the Excel template layout and its other worksheets.
- Name reports using the date or date range found in the CSV data.
- Alert on **Max below 0.1 kW or above 4 kW**, per PDU. Exactly 0.1 and 4 kW are accepted; Average is not checked against these thresholds.
- Detect missing, malformed and impossible calendar dates, including leap-year errors.
- Send the generated Excel to Telegram after successful processing, even when there are no alerts.
- Keep missing rack names and an absent partner PDU explicitly marked as `Неизвестно` ("Unknown"), unless strict options are enabled.

This is a batch application: it processes files when you run it. It does not continuously monitor equipment, retrieve readings through an API, or receive CSV files from Telegram. Numerical values are interpreted as **kW**; there is no automatic conversion from watts.

## Requirements

Choose either:

- **Docker:** Docker Engine on Linux, or Docker Desktop in Linux-container mode on Windows. Python is included in the image.
- **Local Python:** Python 3.10 or newer and the packages in `requirements.txt`. The Docker image uses Python 3.12.

The first Docker build needs internet access to download its base image and dependencies. Telegram delivery needs HTTPS access to `api.telegram.org` and a bot token/chat ID.

Official installation guides: [Docker on Ubuntu](https://docs.docker.com/engine/install/ubuntu/), [Docker Desktop on Windows](https://docs.docker.com/desktop/setup/install/windows-install/).

## Download

On this repository's GitHub page, select **Code → Download ZIP**, extract the archive, and open a terminal inside the extracted project folder.

Alternatively, copy the repository's HTTPS URL from **Code** and clone it:

```bash
git clone YOUR_REPOSITORY_URL powerusage-automation
cd powerusage-automation
```

Replace `YOUR_REPOSITORY_URL` with the actual URL before running the command. [GitHub cloning guide](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository).

## Try the demo with Docker

The `input` folder starts empty. Synthetic demo CSV files are stored separately in `demo/input`.

### Linux

Start Docker, then run:

```bash
bash run.sh --demo
```

The launcher builds the image `powerusage-automation:local` if it is missing, connects the required folders, and runs the application. It uses `sudo docker` when your account cannot access Docker directly.

### Windows PowerShell

Start Docker Desktop and verify that its engine is Linux:

```powershell
docker info --format '{{.OSType}}'
```

Expected output: `linux`.

Run the demo from the extracted project folder:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1 -Demo
```

This execution-policy option applies to that PowerShell process; it does not change your system-wide policy. Use it only for a script you have reviewed and trust.

For both operating systems, the demo produces **three rows** and saves:

```text
output/PowerUsage_2026-10-06.xlsx
```

Telegram is disabled for these commands. Open the report from the local `output` folder, not from an earlier downloaded or cloud copy.

## Process your own CSV files

1. Put your CSV exports directly in `input`. Subfolders are not scanned.
2. Keep `templates/2026_PowerUsageTable_template.xlsx` in place.
3. Run:

Linux:

```bash
bash run.sh
```

Windows:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1
```

Files such as `PowerUsage_2026-10-06.xlsx` or `PowerUsage_2026-10-06_to_2026-10-07.xlsx` appear in `output`.

Successful runs replace the report for the same date/date range. Rows come only from the current batch; the program does not append previous output history. Different date/date-range filenames remain separate. On a processing failure, an older report may still exist: check the exit status and the `Saved ...` line before treating it as a new result. Close the report in a spreadsheet application before rerunning.

After editing Python code or dependencies, rebuild the image:

```bash
bash run.sh --rebuild
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1 -Rebuild
```

Changing CSV files does not require an image rebuild.

## Supported CSV format

Example for a line-A source:

```csv
PDU_2A1;
slot:/Services/ReportService/PDU_2A1;
06.10.2026;
;
Rack_Name;Example rack A11
11A_PDU2_max;11A_PDU2_average
0.35;0.24
```

In this project's naming convention, `PDU_2A1` means floor 2, line A, source/device number 1. **Line A maps to PDU2; line B maps to PDU1.** A line-B counterpart uses metric names such as `11A_PDU1_max` and `11A_PDU1_average`.

`11A`, `A11` and `2F_A11` identify the same rack on floor 2. A file can contain multiple supported rack blocks; the source filename does not determine the number of output rows. This is the export convention supported by this application, not a universal PDU naming standard.

Supported dates: `DD.MM.YYYY` and `YYYY-MM-DD`. Supported encodings: UTF-8 (with or without BOM) and Windows-1251. Supported delimiters: semicolon, comma and tab. Decimal commas are accepted when the CSV values are correctly delimited/quoted.

Average and Max must be finite numbers, and Average cannot exceed Max. Missing or malformed numeric cells are errors, rather than invented zeroes. Identical duplicate readings are merged; conflicting values or rack names are rejected. The optional `Rack_Name` metadata must precede its rack's metric block. Spelling validation against a list of permitted names is not implemented.

The seven report columns are:

| Column | Meaning |
| --- | --- |
| `Rack_No.` | Normalized rack identifier |
| `Rack_Name` | CSV name, or `Неизвестно` |
| `PDU1_avg`, `PDU1_max` | Line-B readings |
| `PDU2_avg`, `PDU2_max` | Line-A readings |
| `Date` | Reporting date from the CSV |

The Excel template supplies layout and formatting. Its sample data rows are replaced; they are not treated as measurement history. Existing content on other worksheets, including `Tariff`, is preserved.

## Telegram setup and delivery

See [the Telegram guide](docs/TELEGRAM.md) for creating a bot and obtaining your chat ID.

Create a **local** `telegram.env` from the blank example and fill in your own values in a text editor. Do not put a token directly in a terminal command or source file.

Linux:

```bash
umask 077
cp -n telegram.env.example telegram.env
chmod 600 telegram.env
nano telegram.env
```

Windows:

```powershell
if (-not (Test-Path -LiteralPath .\telegram.env)) {
    Copy-Item -LiteralPath .\telegram.env.example -Destination .\telegram.env
}
notepad .\telegram.env
```

The two settings are literal `KEY=value` lines, without quotes:

```text
TELEGRAM_BOT_TOKEN=
TELEGRAM_CHAT_ID=
```

Run with Telegram enabled:

```bash
bash run.sh --telegram
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\run.ps1 -Telegram
```

After successful processing, the bot sends the Excel even when there are no warnings. Detected issues are grouped by source:

```text
PDU_2A1:

Шкаф: 11A
Дата: 2026-10-06
Max: 4.2 кВт
Файл: PDU_2A1.csv
Строка: 6
```

Alert labels currently remain in Russian. A syntactically valid date for the wrong measurement day cannot be detected without an external reference date.

`Telegram messages sent: 0` counts **warning messages**, not Excel attachments. Repeated runs send the report and warnings again. On invalid input, date/Max warnings can still be delivered, but a stale Excel is not sent. On delivery failure, the local report remains; requests are not automatically retried because delivery after a timeout may be uncertain.

## Local Python usage

Linux:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python main.py --input-dir demo/input --check --alerts-preview
.venv/bin/python main.py --input-dir demo/input
```

Windows:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py --input-dir demo/input --check --alerts-preview
.\.venv\Scripts\python.exe main.py --input-dir demo/input
```

For local Telegram delivery, add `--telegram --telegram-env telegram.env`.

Other useful options:

| Option | Purpose |
| --- | --- |
| `--input-dir incoming` | Use another CSV folder |
| `--template templates/custom.xlsx` | Use another compatible template |
| `--output output/custom.xlsx` | Choose an explicit filename instead of a date-based one |
| `--check --alerts-preview` | Inspect data/warnings without writing files or sending messages |
| `--require-both-pdus` | Reject incomplete PDU pairs |
| `--rack-name-source strict-csv` | Require a rack name in the CSV |
| `--rack-name-source template` | Explicitly allow fallback names from a trustworthy template |

## Project layout

```text
.
├── main.py                       # CSV parsing and Excel reports
├── telegram_alerts.py            # Messages and document uploads
├── telegram_setup.py             # Find a bot's recent chat IDs
├── container_entrypoint.py       # Protected temporary settings for Docker
├── Dockerfile
├── requirements.txt
├── run.sh                        # Linux launcher
├── run.ps1                       # Windows launcher
├── telegram.env.example          # Blank settings only
├── input/                        # Your CSV files; ignored by Git
├── templates/                    # Excel layout template
├── output/                       # Reports; ignored by Git
├── demo/input/                   # Synthetic example data
├── tests/                        # Offline tests and fixtures
├── docs/                         # Telegram, troubleshooting and publishing guides
└── .github/workflows/ci.yml       # Automated tests and Docker smoke test
```

## Tests

```bash
python -m unittest discover -s tests -v
```

Use your virtual environment's Python executable if needed. Tests mock Telegram requests and do not require a real bot or credentials. GitHub Actions is configured to run tests, check the Linux launcher syntax, build the image and process the demo on push and pull requests. Its first hosted run occurs after publication.

## Security and troubleshooting

Real credentials, private input files, outputs, virtual environments and built image archives are excluded from the source package and `.gitignore`. The Docker build copies only allowlisted code and dependencies. Launchers connect inputs/templates/settings read-only, provide temporary writable `/tmp`, and keep output on the host. A mounted Telegram settings file is copied to an owner-only temporary file inside the container and removed afterward; this handles Windows bind mounts without weakening the application's file-permission validation. Restrict access to the original settings file using Linux permissions or Windows ACLs.

Telegram uses HTTPS with certificate checks and blocked redirects. Error messages do not print request URLs or tokens. These controls do not protect against a compromised host or a Docker administrator who can inspect running containers. See [SECURITY.md](SECURITY.md) and [troubleshooting](docs/TROUBLESHOOTING.md).

For preparing your own GitHub repository or distributing Docker images as Release assets, follow [PUBLISH_TO_GITHUB.md](docs/PUBLISH_TO_GITHUB.md). A license has not been selected yet; the publication guide explains where to add one.
