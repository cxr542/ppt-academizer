#!/usr/bin/env python3
"""Copy ppt-test build engine into apps/ppt-academizer/engine for standalone runs.

WARNING: ``engine/`` is edited directly in this repo (chart copy, classifier, sanitize, ...
fixes since v1.6.6) and the ppt-test copy is OLDER. Syncing replaces ``engine/`` with that
older copy and silently reverts those fixes.

Therefore, when ``engine/`` already exists this script does nothing unless ``--force`` is
given, and even then the current ``engine/`` is moved to ``engine.bak-<timestamp>/`` (never
deleted) so it can be restored.

Usage:
  python scripts/sync_engine_from_ppt_test.py            # fresh checkout without engine/ only
  python scripts/sync_engine_from_ppt_test.py --force    # replace engine/, keep a backup
  PPT_TEST_ROOT=/path/to/ppt-test python scripts/sync_engine_from_ppt_test.py --force
"""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENGINE = ROOT / "engine"

# Monorepo default when PPT_TEST_ROOT unset
DEFAULT_PPT_TEST = ROOT.parent.parent / "cursorstudy" / "experiments" / "ppt-test"


def _ppt_test_root() -> Path:
    import os

    env = os.environ.get("PPT_TEST_ROOT", "").strip()
    if env:
        return Path(env).expanduser().resolve()
    return DEFAULT_PPT_TEST.resolve()


REFUSAL = """\
Refusing to overwrite engine/ (it already exists).

engine/ is maintained in THIS repo and contains fixes that the ppt-test copy does not have
(chart part copy, table/cover classification, SVG-only graphics, content-loss warnings, ...).
Syncing would revert them.

If you really want to replace it:
  python scripts/sync_engine_from_ppt_test.py --force
The current engine/ is then moved to engine.bak-<timestamp>/ (not deleted); diff it against
the new engine/ and re-apply what you need.
"""


def _backup_engine() -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = ROOT / f"engine.bak-{stamp}"
    n = 1
    while backup.exists():
        n += 1
        backup = ROOT / f"engine.bak-{stamp}-{n}"
    ENGINE.rename(backup)
    return backup


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Copy the ppt-test engine into engine/.")
    parser.add_argument(
        "--force",
        action="store_true",
        help="replace an existing engine/ (it is moved to engine.bak-<timestamp>/ first)",
    )
    args = parser.parse_args(argv)

    src = _ppt_test_root()
    scripts_src = src / "scripts"
    examples_src = src / "docs" / "examples"
    office_src = src / ".cursor" / "skills" / "pptx" / "scripts" / "office"
    if not scripts_src.is_dir():
        print(f"Missing ppt-test scripts: {scripts_src}", file=sys.stderr)
        return 2

    dst_scripts = ENGINE / "scripts"
    dst_examples = ENGINE / "docs" / "examples"
    dst_office = ENGINE / "office"
    if ENGINE.exists():
        if not args.force:
            print(REFUSAL, file=sys.stderr)
            return 3
        backup = _backup_engine()
        print(f"Moved existing engine/ to {backup}")
    dst_scripts.mkdir(parents=True)
    dst_examples.mkdir(parents=True)

    for py in sorted(scripts_src.glob("*.py")):
        shutil.copy2(py, dst_scripts / py.name)

    if examples_src.is_dir():
        for j in sorted(examples_src.glob("*.json")):
            shutil.copy2(j, dst_examples / j.name)

    if office_src.is_dir():
        shutil.copytree(office_src, dst_office)
    else:
        print(f"WARN: missing OOXML office bundle: {office_src}", file=sys.stderr)

  # Engine root marker (version from migrate_version.py)
    ver_src = scripts_src / "migrate_version.py"
    ver_text = ver_src.read_text(encoding="utf-8") if ver_src.is_file() else ""
    engine_ver = "unknown"
    for line in ver_text.splitlines():
        if line.strip().startswith("MIGRATE_ENGINE_VERSION"):
            engine_ver = line.split("=", 1)[1].strip().strip('"').strip("'")
            break
    (ENGINE / "ENGINE_VERSION").write_text(
        f"ppt-engine={engine_ver}\nsource={src}\n",
        encoding="utf-8",
    )
    (ENGINE / "README.md").write_text(
        "# Bundled PPT build engine\n\n"
        "Originally synced from ppt-test. Since v1.6.6 this directory is **edited directly in\n"
        "this repo** and is newer than ppt-test. Re-running the sync reverts those fixes, so\n"
        "`scripts/sync_engine_from_ppt_test.py` refuses to run when `engine/` exists unless\n"
        "`--force` is given (and then keeps `engine.bak-<timestamp>/`).\n\n"
        f"Current engine version: **{engine_ver}**\n",
        encoding="utf-8",
    )

    print(f"Synced engine from {src}")
    print(f"  scripts: {len(list(dst_scripts.glob('*.py')))} files")
    print(f"  examples: {len(list(dst_examples.glob('*.json')))} json")
    if dst_office.is_dir():
        print(f"  office: validators + schemas ({dst_office})")
    print(f"  -> {ENGINE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
