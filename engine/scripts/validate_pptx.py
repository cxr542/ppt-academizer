#!/usr/bin/env python3
"""Run OOXML validator on a .pptx file (bundled engine/office or legacy ppt-test path)."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _resolve_office_dir() -> Path:
    bundled = ROOT / "office"
    if (bundled / "validate.py").is_file():
        return bundled
    legacy = ROOT / ".cursor" / "skills" / "pptx" / "scripts" / "office"
    if (legacy / "validate.py").is_file():
        return legacy
    raise FileNotFoundError(
        "office/validate.py not found. Run: python scripts/sync_engine_from_ppt_test.py"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate .pptx OOXML")
    parser.add_argument("pptx", type=Path, help="Path to .pptx file")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument(
        "--auto-repair",
        action="store_true",
        help="Run validator auto-repair before validation",
    )
    args = parser.parse_args()

    pptx = args.pptx.resolve()
    if not pptx.is_file() or pptx.suffix.lower() != ".pptx":
        print(f"Error: not a .pptx file: {pptx}", file=sys.stderr)
        sys.exit(2)

    office = _resolve_office_dir()
    validate_py = office / "validate.py"

    cmd = [sys.executable, str(validate_py), str(pptx)]
    if args.verbose:
        cmd.append("-v")
    if args.auto_repair:
        cmd.append("--auto-repair")

    r = subprocess.run(cmd, cwd=str(office))
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()
