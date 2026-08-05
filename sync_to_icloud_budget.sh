#!/usr/bin/env bash
# Copy this repo's budget files into the local iCloud budget folder (Mac only).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
DEST="${BUDGET_ICLOUD_DIR:-/Users/msanes/Library/Mobile Documents/com~apple~CloudDocs/Cloud/Cursor/budget}"

if [[ ! -d "$DEST" ]]; then
  echo "ERROR: iCloud budget folder not found:"
  echo "  $DEST"
  echo "Run this on the Mac where that path exists, or set BUDGET_ICLOUD_DIR."
  exit 1
fi

# Mirror key trees into the iCloud clone/folder (no .git overwrite).
rsync -a --delete \
  --exclude '.git/' \
  --exclude '.cursor/' \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  "$ROOT/AVA/" "$DEST/AVA/"
rsync -a \
  --exclude '.git/' \
  --exclude '__pycache__/' \
  "$ROOT/BONORINO/" "$DEST/BONORINO/" 2>/dev/null || true
rsync -a \
  --exclude '.git/' \
  --exclude '__pycache__/' \
  "$ROOT/CIUDAD/" "$DEST/CIUDAD/"
rsync -a \
  --exclude '.git/' \
  --exclude '__pycache__/' \
  "$ROOT/OHIGGINS/" "$DEST/OHIGGINS/" 2>/dev/null || true
rsync -a \
  --exclude '.git/' \
  --exclude '__pycache__/' \
  "$ROOT/SOL/" "$DEST/SOL/"

if [[ -f "$ROOT/README.md" ]]; then
  cp -a "$ROOT/README.md" "$DEST/README.md"
fi
if [[ -f "$ROOT/sync_to_icloud_budget.sh" ]]; then
  cp -a "$ROOT/sync_to_icloud_budget.sh" "$DEST/sync_to_icloud_budget.sh"
fi

echo "Synced to: $DEST"
ls -la "$DEST/AVA/Gastos_AVA_2026.xlsx" "$DEST/CIUDAD/Gastos_CIUDAD_1132_2026.xlsx" "$DEST/SOL/Gastos_SOL_2026.xlsx"
