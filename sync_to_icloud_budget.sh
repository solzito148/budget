#!/usr/bin/env bash
# Siempre espejar workspace → local iCloud budget (Mac).
# Uso: ./sync_to_icloud_budget.sh
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
DEST="${BUDGET_ICLOUD_DIR:-/Users/msanes/Library/Mobile Documents/com~apple~CloudDocs/Cloud/Cursor/budget}"

if [[ "$ROOT" -ef "$DEST" ]]; then
  echo "Workspace already is iCloud budget folder: $DEST"
  echo "Nothing to mirror."
  exit 0
fi

if [[ ! -d "$DEST" ]]; then
  echo "ERROR: iCloud budget folder not found:"
  echo "  $DEST"
  echo "Run on the Mac where that path exists, or set BUDGET_ICLOUD_DIR."
  exit 1
fi

mkdir -p "$DEST"

rsync_tree() {
  local src="$1" dst="$2"
  [[ -d "$src" ]] || return 0
  mkdir -p "$dst"
  rsync -a --delete \
    --exclude '.git/' \
    --exclude '__pycache__/' \
    --exclude '*.pyc' \
    --exclude '.DS_Store' \
    "$src/" "$dst/"
}

rsync_tree "$ROOT/AVA" "$DEST/AVA"
rsync_tree "$ROOT/BONORINO" "$DEST/BONORINO"
rsync_tree "$ROOT/CIUDAD" "$DEST/CIUDAD"
rsync_tree "$ROOT/OHIGGINS" "$DEST/OHIGGINS"
rsync_tree "$ROOT/SOL" "$DEST/SOL"
rsync_tree "$ROOT/.cursor/rules" "$DEST/.cursor/rules"

cp -a "$ROOT/sync_to_icloud_budget.sh" "$DEST/sync_to_icloud_budget.sh"
[[ -f "$ROOT/README.md" ]] && cp -a "$ROOT/README.md" "$DEST/README.md"

echo "Synced workspace → iCloud:"
echo "  $DEST"
ls -la "$DEST/AVA/Gastos_AVA_2026.xlsx" \
       "$DEST/CIUDAD/Gastos_CIUDAD_1132_2026.xlsx" \
       "$DEST/SOL/Gastos_SOL_2026.xlsx"
