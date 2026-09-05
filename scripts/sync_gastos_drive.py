#!/usr/bin/env python3
"""Sync property expense workbooks into Gastos/ and optionally upload to Google Drive.

Local mirror (always):
  CIUDAD/BONORINO/OHIGGINS/SOL → Gastos/

Remote (optional, needs credentials):
  Drive folder 1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x

Auth options (first match wins):
  1. GOOGLE_APPLICATION_CREDENTIALS → service account JSON
  2. Gastos/.drive_token.json → OAuth user token from desktop auth
  3. --dry-run → only refresh local Gastos/ copies
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GASTOS = ROOT / "Gastos"
DRIVE_FOLDER_ID = "1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x"
DRIVE_URL = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"

FILES = [
    ROOT / "CIUDAD" / "Gastos_CIUDAD_1132_2026.xlsx",
    ROOT / "BONORINO" / "Gastos_BONORINO_2026.xlsx",
    ROOT / "OHIGGINS" / "Gastos_OHIGGINS_2026.xlsx",
    ROOT / "SOL" / "Gastos_SOL_2026.xlsx",
]


def refresh_local() -> list[Path]:
    GASTOS.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for src in FILES:
        if not src.exists():
            raise FileNotFoundError(f"Missing source workbook: {src}")
        dest = GASTOS / src.name
        shutil.copy2(src, dest)
        copied.append(dest)
        print(f"local  {src.relative_to(ROOT)} → Gastos/{src.name} ({dest.stat().st_size} bytes)")
    return copied


def upload_drive(paths: list[Path]) -> None:
    try:
        from google.oauth2 import service_account
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError as exc:
        raise SystemExit(
            "Missing Google API libs. Install with:\n"
            "  pip install google-api-python-client google-auth google-auth-oauthlib\n"
            f"Local Gastos/ copies are ready; upload to {DRIVE_URL} manually or after auth."
        ) from exc

    scopes = ["https://www.googleapis.com/auth/drive.file"]
    creds = None
    sa_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    token_path = GASTOS / ".drive_token.json"

    if sa_path and Path(sa_path).exists():
        creds = service_account.Credentials.from_service_account_file(sa_path, scopes=scopes)
        print(f"auth   service account ({sa_path})")
    elif token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), scopes=scopes)
        print(f"auth   oauth token ({token_path})")
    else:
        raise SystemExit(
            "No Google Drive credentials in this environment.\n"
            f"Local copies are in Gastos/. Target folder: {DRIVE_URL}\n"
            "Provide GOOGLE_APPLICATION_CREDENTIALS or Gastos/.drive_token.json, "
            "or authenticate Composio Google Drive in Cursor desktop."
        )

    service = build("drive", "v3", credentials=creds, cache_discovery=False)

    # Map existing files in folder by name
    existing: dict[str, str] = {}
    page_token = None
    query = f"'{DRIVE_FOLDER_ID}' in parents and trashed=false"
    while True:
        resp = (
            service.files()
            .list(
                q=query,
                spaces="drive",
                fields="nextPageToken, files(id, name)",
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        for item in resp.get("files", []):
            existing[item["name"]] = item["id"]
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    for path in paths:
        media = MediaFileUpload(
            str(path),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        )
        if path.name in existing:
            service.files().update(
                fileId=existing[path.name],
                media_body=media,
                supportsAllDrives=True,
            ).execute()
            print(f"drive  updated {path.name} ({existing[path.name]})")
        else:
            meta = {"name": path.name, "parents": [DRIVE_FOLDER_ID]}
            created = (
                service.files()
                .create(body=meta, media_body=media, fields="id,name", supportsAllDrives=True)
                .execute()
            )
            print(f"drive  created {created['name']} ({created['id']})")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only refresh local Gastos/ copies; do not upload",
    )
    parser.add_argument(
        "--upload",
        action="store_true",
        help="Upload to Google Drive after refreshing local copies",
    )
    args = parser.parse_args()
    paths = refresh_local()
    if args.dry_run and not args.upload:
        print(f"dry-run complete → {DRIVE_URL}")
        return 0
    if args.upload or not args.dry_run:
        # Default behavior: try upload; if no creds, local-only with clear message
        try:
            upload_drive(paths)
        except SystemExit as exc:
            if args.upload:
                raise
            print(str(exc), file=sys.stderr)
            return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
