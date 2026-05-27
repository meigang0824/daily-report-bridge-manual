#!/bin/zsh
set -euo pipefail

bundle_dir="$(cd "$(dirname "$0")" && pwd)"
env_file="${1:-"$bundle_dir/.env"}"

if [[ ! -f "$env_file" ]]; then
  echo "Missing configuration file: $env_file"
  echo "Copy .env.example to .env, fill in database credentials, then run again."
  exit 1
fi

set -a
source "$env_file"
set +a

exec python3 "$bundle_dir/bridge/operations_data_bridge.py"
