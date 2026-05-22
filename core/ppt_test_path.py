"""Resolve PPT build engine root (bundled engine/ or monorepo ppt-test)."""

from __future__ import annotations

import os
from pathlib import Path


def resolve_ppt_test_root() -> Path:
    """Engine root for `from scripts.*` imports.

    Priority:
    1. ``PPT_ENGINE_ROOT`` or ``PPT_TEST_ROOT`` (explicit override)
    2. ``apps/ppt-academizer/engine`` (bundled copy — standalone)
    3. ``cursorstudy/experiments/ppt-test`` (monorepo sibling)
    """
    for key in ("PPT_ENGINE_ROOT", "PPT_TEST_ROOT"):
        env = os.environ.get(key, "").strip()
        if env:
            return Path(env).expanduser().resolve()

    academizer = Path(__file__).resolve().parent.parent
    bundled = academizer / "engine"
    if (bundled / "scripts" / "academy_deck_build_lib.py").is_file():
        return bundled.resolve()

    return (
        academizer.parent.parent / "cursorstudy" / "experiments" / "ppt-test"
    ).resolve()


def ensure_engine_on_path() -> Path:
    """Insert engine root on sys.path if missing; return engine root."""
    import sys

    root = resolve_ppt_test_root()
    s = str(root)
    if s not in sys.path:
        sys.path.insert(0, s)
    return root
