#!/usr/bin/env python3
"""Post-process an academy deck: content title (v9) + open in slide view."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.academy_deck_build_lib import save_academy_deck  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Polish academy PPTX (title + sldView)")
    parser.add_argument("pptx", type=Path, help="Path to .pptx")
    args = parser.parse_args()
    path = args.pptx.resolve()
    if not path.is_file():
        print(f"Not found: {path}", file=sys.stderr)
        return 2
    prs = Presentation(str(path))
    save_academy_deck(prs, path)
    print(f"Polished: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
