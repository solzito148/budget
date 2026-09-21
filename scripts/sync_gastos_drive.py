#!/usr/bin/env python3
"""Mirror property Excel workbooks into Gastos/ and sync to Google Drive/Sheets.

Cloud-agent canonical (editable via Composio Google Sheets):
  https://docs.google.com/spreadsheets/d/14398lhBiIOkSAXSbwTp_XTF4M_HmlpTCWkWFCfJIrto/edit
  Spreadsheet ID: 14398lhBiIOkSAXSbwTp_XTF4M_HmlpTCWkWFCfJIrto

Drive folder for xlsx mirrors:
  https://drive.google.com/drive/folders/1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x
  Folder ID: 1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x

Local Mac iCloud remains the personal copy:
  ~/Library/Mobile Documents/com~apple~CloudDocs/GastosCiudad/
  Use scripts/sync_gastos_icloud.py on the Mac.

Usage:
  python3 scripts/sync_gastos_drive.py              # refresh Gastos/ only
  python3 scripts/sync_gastos_drive.py --print-plan # show Composio upload plan
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GASTOS = ROOT / "Gastos"
DRIVE_FOLDER_ID = "1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x"
DRIVE_URL = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"
SHEETS_ID = "14398lhBiIOkSAXSbwTp_XTF4M_HmlpTCWkWFCfJIrto"
SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/edit"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

FILES = [
    ROOT / "BONORINO" / "Gastos_BONORINO_2026.xlsx",
    ROOT / "CIUDAD" / "Gastos_CIUDAD_1132_2026.xlsx",
    ROOT / "OHIGGINS" / "Gastos_OHIGGINS_2026.xlsx",
    ROOT / "SOL" / "Gastos_SOL_2026.xlsx",
]


def refresh_local() -> list[Path]:
    GASTOS.mkdir(parents=True, exist_ok=True)
    out: list[Path] = []
    for src in FILES:
        if not src.exists():
            print(f"skip   missing {src.relative_to(ROOT)}", file=sys.stderr)
            continue
        dest = GASTOS / src.name
        shutil.copy2(src, dest)
        out.append(dest)
        print(f"local  {src.relative_to(ROOT)} → Gastos/{src.name} ({dest.stat().st_size} bytes)")
    return out


def print_composio_plan(paths: list[Path]) -> None:
    print("\n# Composio Google Drive + Sheets sync (after OAuth is ACTIVE)")
    print(f"# Canonical Sheet: {SHEETS_URL}")
    print(f"# Spreadsheet ID: {SHEETS_ID}")
    print(f"# Drive folder:   {DRIVE_URL}")
    print("# Steps for the agent:")
    print("1. COMPOSIO_MANAGE_CONNECTIONS toolkit=googledrive + googlesheets (must be ACTIVE)")
    print("2. GOOGLESHEETS_GET_SHEET_NAMES spreadsheet_id=%s" % SHEETS_ID)
    print("3. Read/update monthly expenses via GOOGLESHEETS_VALUES_GET / VALUES_UPDATE")
    print("4. GOOGLEDRIVE_FIND_FILE folder_id=%s  q=\"trashed = false\"" % DRIVE_FOLDER_ID)
    print("5. For each workbook mirror: upload via GOOGLEDRIVE_CREATE_FILE (new) or")
    print("   GOOGLEDRIVE_UPLOAD_UPDATE_FILE (existing fileId) with file_to_upload s3key")
    print("   from COMPOSIO_REMOTE_WORKBENCH upload_local_file(path).")
    print("\ncanonical_sheet:")
    print(json.dumps({"spreadsheet_id": SHEETS_ID, "url": SHEETS_URL}))
    print("\nfiles:")
    for p in paths:
        print(
            json.dumps(
                {
                    "name": p.name,
                    "path": str(p),
                    "mimetype": XLSX,
                    "parents": [DRIVE_FOLDER_ID],
                }
            )
        )


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--print-plan", action="store_true", help="Print Composio upload plan")
    args = ap.parse_args()
    paths = refresh_local()
    if not paths:
        return 1
    print(f"done   Gastos/ mirror ready ({len(paths)} files)")
    print(f"sheet  {SHEETS_URL}")
    print(f"drive  {DRIVE_URL}")
    if args.print_plan:
        print_composio_plan(paths)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
