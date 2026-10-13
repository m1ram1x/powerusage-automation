param(
    [switch]$Demo,
    [switch]$Telegram,
    [switch]$Rebuild
)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    throw 'Install and start Docker Desktop first.'
}
$engineType = & docker info --format '{{.OSType}}'
if ($LASTEXITCODE -ne 0 -or $engineType -ne 'linux') {
    throw 'Start Docker Desktop in Linux-container mode.'
}
foreach ($folder in @('input', 'output')) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}
$inputPath = Join-Path $PSScriptRoot 'input'
if ($Demo) { $inputPath = Join-Path $PSScriptRoot 'demo/input' }
if (-not (Test-Path -LiteralPath 'templates/2026_PowerUsageTable_template.xlsx')) {
    throw 'The Excel template is missing from templates/.'
}
$csvFiles = @(Get-ChildItem -LiteralPath $inputPath -File | Where-Object { $_.Extension -ieq '.csv' })
if ($csvFiles.Count -eq 0) {
    throw 'Add CSV files to input/, or run this script with -Demo.'
}
$image = 'powerusage-automation:local'
& docker image inspect $image *> $null
$imageExists = $LASTEXITCODE -eq 0
if ($Rebuild -or -not $imageExists) {
    & docker build -t $image .
    if ($LASTEXITCODE -ne 0) { throw 'Docker image build failed.' }
}
$dockerArguments = @(
    'run', '--rm', '--user', '10001:10001',
    '--read-only', '--cap-drop', 'ALL',
    '--security-opt', 'no-new-privileges',
    '--tmpfs', '/tmp:rw,noexec,nosuid,nodev,size=128m,mode=1777',
    '--mount', "type=bind,source=$inputPath,target=/app/input,readonly",
    '--mount', "type=bind,source=$PSScriptRoot/templates,target=/app/templates,readonly",
    '--mount', "type=bind,source=$PSScriptRoot/output,target=/app/output"
)
if ($Telegram) {
    if (-not (Test-Path -LiteralPath 'telegram.env' -PathType Leaf)) {
        throw 'Create local telegram.env using telegram.env.example first.'
    }
    $dockerArguments += @('--mount', "type=bind,source=$PSScriptRoot/telegram.env,target=/run/secrets/telegram.env,readonly")
}
$dockerArguments += @($image, '--template', 'templates/2026_PowerUsageTable_template.xlsx')
if ($Telegram) { $dockerArguments += @('--telegram', '--telegram-env', '/run/secrets/telegram.env') }
& docker @dockerArguments
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Host "Reports saved in: $PSScriptRoot/output"
