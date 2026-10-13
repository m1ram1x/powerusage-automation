#!/usr/bin/env bash
set -euo pipefail
cd -- "$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"

task_telegram=0
task_demo=0
task_rebuild=0
task_image=powerusage-automation:local
while [[ $# -gt 0 ]]; do
  case "$1" in
    --telegram) task_telegram=1; shift ;;
    --demo) task_demo=1; shift ;;
    --rebuild) task_rebuild=1; shift ;;
    --help)
      printf 'Usage: bash run.sh [--demo] [--telegram] [--rebuild]\n'
      exit 0 ;;
    *) printf 'Unknown option: %s\n' "$1" >&2; exit 2 ;;
  esac
done
command -v docker >/dev/null 2>&1 || { printf 'Install Docker first.\n' >&2; exit 1; }
task_docker=(docker)
if ! docker info >/dev/null 2>&1; then task_docker=(sudo docker); fi
task_engine="$("${task_docker[@]}" info --format '{{.OSType}}')"
[[ "$task_engine" == linux ]] || { printf 'A running Linux Docker engine is required.\n' >&2; exit 1; }
task_input="$PWD/input"
if [[ "$task_demo" == 1 ]]; then task_input="$PWD/demo/input"; fi
mkdir -p input output
[[ -f templates/2026_PowerUsageTable_template.xlsx ]] || { printf 'Excel template is missing from templates/.\n' >&2; exit 1; }
task_found=0
while IFS= read -r -d '' task_csv; do task_found=1; break; done < <(find "$task_input" -maxdepth 1 -type f -iname '*.csv' -print0)
[[ "$task_found" == 1 ]] || { printf 'Add CSV files to input/, or try: bash run.sh --demo\n' >&2; exit 1; }
if [[ "$task_rebuild" == 1 ]] || ! "${task_docker[@]}" image inspect "$task_image" >/dev/null 2>&1; then
  "${task_docker[@]}" build -t "$task_image" .
fi
task_mounts=(
  --mount "type=bind,source=$task_input,target=/app/input,readonly"
  --mount "type=bind,source=$PWD/templates,target=/app/templates,readonly"
  --mount "type=bind,source=$PWD/output,target=/app/output"
)
task_args=(--template templates/2026_PowerUsageTable_template.xlsx)
if [[ "$task_telegram" == 1 ]]; then
  [[ -f telegram.env ]] || { printf 'Create local telegram.env using telegram.env.example first.\n' >&2; exit 1; }
  chmod 600 telegram.env
  task_mounts+=(--mount "type=bind,source=$PWD/telegram.env,target=/run/secrets/telegram.env,readonly")
  task_args+=(--telegram --telegram-env /run/secrets/telegram.env)
fi
"${task_docker[@]}" run --rm \
  --user "$(id -u):$(id -g)" \
  --read-only --cap-drop ALL --security-opt no-new-privileges \
  --tmpfs /tmp:rw,noexec,nosuid,nodev,size=128m,mode=1777 \
  "${task_mounts[@]}" "$task_image" "${task_args[@]}"
printf 'Reports saved in: %s/output\n' "$PWD"
