#!/usr/bin/env python3
"""Copy partner deck fixtures from manifest paths (local only)."""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "tests" / "fixtures" / "manifest.json"
FIXTURES = ROOT / "tests" / "fixtures"


def _expand(p: str) -> Path:
    return Path(p).expanduser()


def _resolve_entry(entry: dict) -> Path | None:
    paths = entry.get("paths") or []
    glob_pat = entry.get("glob")
    for raw in paths:
        base = _expand(raw)
        if glob_pat and base.is_dir():
            matches = sorted(base.glob(glob_pat))
            if matches:
                return matches[0]
        if base.is_file():
            return base
    return None


def main() -> int:
    if not MANIFEST.is_file():
        print(f"Missing {MANIFEST}", file=sys.stderr)
        return 2

    data = json.loads(MANIFEST.read_text(encoding="utf-8"))
    FIXTURES.mkdir(parents=True, exist_ok=True)
    ok, skip, fail = 0, 0, 0

    for entry in data.get("sources", []):
        target = FIXTURES / entry["target"]
        src = _resolve_entry(entry)
        optional = entry.get("optional", False)
        if src is None:
            if optional:
                skip += 1
                print(f"SKIP (optional) {entry['id']}: no source")
            else:
                fail += 1
                print(f"MISSING {entry['id']}: not found in manifest paths")
            continue
        shutil.copy2(src, target)
        ok += 1
        print(f"OK {entry['id']}: {src} -> {target}")

    print(f"\n{ok} copied, {skip} optional skipped, {fail} missing")
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
