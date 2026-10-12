"""Resolve the PPT build engine root (the bundled ``engine/`` directory).

The module keeps its old name because many callers import it; the former ``ppt-test`` project
is retired and is no longer used as a fallback (its copy was outdated and silently produced
conversions without the fixes made in ``engine/``).
"""

from __future__ import annotations

import os
from pathlib import Path

_ENGINE_MARKER = Path("scripts") / "academy_deck_build_lib.py"


def resolve_ppt_test_root(app_root: Path | None = None) -> Path:
    """Engine root for `from scripts.*` imports.

    Priority:
    1. ``PPT_ENGINE_ROOT`` or ``PPT_TEST_ROOT`` (explicit override)
    2. ``<app root>/engine`` (the bundled engine, maintained in this repo)

    Raises ``FileNotFoundError`` when neither exists, instead of falling back to a stale copy.
    """
    for key in ("PPT_ENGINE_ROOT", "PPT_TEST_ROOT"):
        env = os.environ.get(key, "").strip()
        if env:
            return Path(env).expanduser().resolve()

    root = app_root if app_root is not None else Path(__file__).resolve().parent.parent
    bundled = root / "engine"
    if (bundled / _ENGINE_MARKER).is_file():
        return bundled.resolve()

    raise FileNotFoundError(
        f"Bundled engine not found at {bundled}. Restore it from git "
        "(git checkout -- engine) or point PPT_ENGINE_ROOT at an engine directory."
    )


resolve_engine_root = resolve_ppt_test_root


def ensure_engine_on_path() -> Path:
    """Insert engine root on sys.path if missing; return engine root."""
    import sys

    root = resolve_ppt_test_root()
    s = str(root)
    if s not in sys.path:
        sys.path.insert(0, s)
    return root
