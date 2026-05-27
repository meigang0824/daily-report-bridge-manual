#!/bin/zsh
set -euo pipefail

bundle_dir="$(cd "$(dirname "$0")" && pwd)"
env_file="${1:-"$bundle_dir/.env"}"

if [[ ! -f "$env_file" ]]; then
  echo "Missing configuration file: $env_file"
  echo "Copy .env.example to .env, fill in credentials, then run again."
  exit 1
fi

set -a
source "$env_file"
set +a

export DAILY_LEGACY_MAIN="${DAILY_LEGACY_MAIN/__BUNDLE_DIR__/$bundle_dir}"
export DAILY_RUNTIME_DIR="${DAILY_RUNTIME_DIR/__BUNDLE_DIR__/$bundle_dir}"
export DAILY_FIXTURE_PATH="${DAILY_FIXTURE_PATH/__BUNDLE_DIR__/$bundle_dir}"

mkdir -p "$DAILY_RUNTIME_DIR"
exec python3 "$bundle_dir/bridge/daily_report_bridge.py"
