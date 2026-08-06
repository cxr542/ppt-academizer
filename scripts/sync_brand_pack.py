#!/usr/bin/env python3
"""Sync brand/ from ppt-academizer-assets (or rebuild icons from elements PPTX)."""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--assets",
        type=Path,
        default=Path(__file__).resolve().parents[2] / "ppt-academizer-assets",
        help="Clone of cxr542/ppt-academizer-assets",
    )
    ap.add_argument(
        "--dest",
        type=Path,
        default=Path.home() / ".config" / "ppt-academizer-render" / "brand",
        help="Local brand dir used next to TEMPLATE_PPTX",
    )
    args = ap.parse_args()
    src = args.assets / "brand"
    if not src.is_dir():
        print(f"missing brand pack: {src}", file=sys.stderr)
        return 1
    args.dest.parent.mkdir(parents=True, exist_ok=True)
    if args.dest.exists():
        shutil.rmtree(args.dest)
    shutil.copytree(src, args.dest)
    print(f"synced {src} -> {args.dest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
