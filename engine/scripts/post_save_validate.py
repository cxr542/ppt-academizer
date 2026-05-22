"""After-save PPTX OOXML validation (Anthropic pptx skill)."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def validate_pptx_or_exit(pptx: Path) -> None:
    """Run scripts/validate_pptx.py; exit non-zero on failure.

    Assumes path like ``<repo>/output/name.pptx`` so repo root is ``pptx.parents[1]``.
    """
    pptx = pptx.resolve()
    root = pptx.parents[1]
    script = root / "scripts" / "validate_pptx.py"
    if not script.is_file():
        print(f"validate_pptx.py missing, skip: {script}", file=sys.stderr)
        return
    r = subprocess.run(
        [sys.executable, str(script), str(pptx)],
        cwd=str(root),
    )
    if r.returncode != 0:
        raise SystemExit(r.returncode)
    print("OOXML validation: PASSED")
