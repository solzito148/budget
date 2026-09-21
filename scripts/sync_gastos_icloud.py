#!/usr/bin/env python3
"""Sync property expense workbooks between the git repo and iCloud GastosCiudad.

Canonical (Mac / iCloud):
  /Users/sol/Library/Mobile Documents/com~apple~CloudDocs/GastosCiudad/
    BONORINO/Gastos_BONORINO_2026.xlsx   ← updated every month automatically
    CIUDAD/Gastos_CIUDAD_1132_2026.xlsx
    OHIGGINS/Gastos_OHIGGINS_2026.xlsx
    SOL/Gastos_SOL_2026.xlsx
    AVA/… (optional)

Repo working copies:
  BONORINO/, CIUDAD/, OHIGGINS/, SOL/, AVA/

Usage:
  python3 scripts/sync_gastos_icloud.py           # repo → iCloud
  python3 scripts/sync_gastos_icloud.py --pull    # iCloud → repo
  python3 scripts/sync_gastos_icloud.py --dry-run
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

DEFAULT_ICLOUD = Path(
    "/Users/sol/Library/Mobile Documents/com~apple~CloudDocs/GastosCiudad"
)

# (repo relative path, iCloud relative path)
PAIRS = [
    ("BONORINO/Gastos_BONORINO_2026.xlsx", "BONORINO/Gastos_BONORINO_2026.xlsx"),
    ("CIUDAD/Gastos_CIUDAD_1132_2026.xlsx", "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx"),
    ("OHIGGINS/Gastos_OHIGGINS_2026.xlsx", "OHIGGINS/Gastos_OHIGGINS_2026.xlsx"),
    ("SOL/Gastos_SOL_2026.xlsx", "SOL/Gastos_SOL_2026.xlsx"),
    ("AVA/Gastos_AVA_2026.xlsx", "AVA/Gastos_AVA_2026.xlsx"),
]


def icloud_root() -> Path:
    env = os.environ.get("GASTOS_ICLOUD_ROOT")
    return Path(env) if env else DEFAULT_ICLOUD


def copy_one(src: Path, dest: Path, *, dry_run: bool) -> None:
    if not src.exists():
        print(f"skip   missing source: {src}")
        return
    print(f"{'dry' if dry_run else 'copy'}  {src} → {dest} ({src.stat().st_size} bytes)")
    if dry_run:
        return
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--pull",
        action="store_true",
        help="iCloud → repo (default is repo → iCloud)",
    )
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument(
        "--only",
        choices=["BONORINO", "CIUDAD", "OHIGGINS", "SOL", "AVA"],
        help="Sync a single property",
    )
    args = ap.parse_args()

    base = icloud_root()
    if not base.exists():
        print(
            f"iCloud root not found: {base}\n"
            "This environment cannot see Sol's Mac iCloud folder.\n"
            "On the Mac, run this script after merging/pulling the repo so the\n"
            "canonical GastosCiudad Excel files stay up to date.\n"
            "Override with GASTOS_ICLOUD_ROOT if the path differs.",
            file=sys.stderr,
        )
        return 2

    pairs = PAIRS
    if args.only:
        pairs = [p for p in PAIRS if p[0].startswith(args.only + "/")]

    for repo_rel, icloud_rel in pairs:
        repo_path = ROOT / repo_rel
        icloud_path = base / icloud_rel
        if args.pull:
            copy_one(icloud_path, repo_path, dry_run=args.dry_run)
        else:
            copy_one(repo_path, icloud_path, dry_run=args.dry_run)

    direction = "iCloud → repo" if args.pull else "repo → iCloud"
    print(f"done   {direction} ({len(pairs)} file(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
