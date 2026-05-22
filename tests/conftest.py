"""pytest: bundled engine + app root on sys.path."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.ppt_test_path import ensure_engine_on_path  # noqa: E402

ensure_engine_on_path()
