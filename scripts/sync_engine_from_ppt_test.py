#!/usr/bin/env python3
"""Copy ppt-test build engine into apps/ppt-academizer/engine for standalone runs.

Usage:
  python scripts/sync_engine_from_ppt_test.py
  PPT_TEST_ROOT=/path/to/ppt-test python scripts/sync_engine_from_ppt_test.py
"""

from __future__ import annotations

import shutil
import sys
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


def main() -> int:
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
        shutil.rmtree(ENGINE)
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
        "Auto-synced from ppt-test. Do not edit by hand — run:\n\n"
        "```bash\n"
        "python scripts/sync_engine_from_ppt_test.py\n"
        "```\n\n"
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
