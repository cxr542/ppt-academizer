"""Unique output path helpers — never overwrite an existing deliverable."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path


def unique_output_path(
    directory: Path | str,
    stem: str,
    *,
    suffix: str = ".pptx",
    timestamp: str | None = None,
) -> Path:
    """Return ``directory/stem-YYYYMMDD-HHMMSS[.n].suffix`` that does not exist yet."""
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    stamp = timestamp or datetime.now().strftime("%Y%m%d-%H%M%S")
    safe_stem = (stem or "academy-output").strip().replace("/", "-").replace("\\", "-")
    candidate = directory / f"{safe_stem}-{stamp}{suffix}"
    if not candidate.exists():
        return candidate
    n = 2
    while True:
        candidate = directory / f"{safe_stem}-{stamp}-{n}{suffix}"
        if not candidate.exists():
            return candidate
        n += 1
