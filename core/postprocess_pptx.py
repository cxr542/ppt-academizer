"""Post-save fixes for academy .pptx (Mac PowerPoint open without Repair dialog)."""

from __future__ import annotations

import os
import platform
from pathlib import Path


def maybe_mac_powerpoint_repair(path: Path) -> list[dict]:
    """Run Microsoft PowerPoint repair pass on macOS (accept Repair, save in place).

    OOXML whitespace repair alone does not fix SVG/media packaging that triggers
    the Mac「복구」dialog on partner (§7) decks.
    """
    if platform.system() != "Darwin":
        return []
    if os.environ.get("PPT_ACADEMIZER_SKIP_PP_REPAIR", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return []

    try:
        from scripts.powerpoint_repair_mac import powerpoint_repair_and_save
    except ImportError:
        return [
            {
                "code": "PP_REPAIR_SKIPPED",
                "message": "powerpoint_repair_mac unavailable",
            }
        ]

    path = Path(path).resolve()
    repaired = path.with_suffix(".pp-repair.pptx")
    try:
        powerpoint_repair_and_save(path, repaired, timeout_s=120)
        repaired.replace(path)
        return [{"code": "PP_REPAIR_APPLIED", "message": ""}]
    except Exception as exc:
        if repaired.is_file():
            repaired.unlink(missing_ok=True)
        return [{"code": "PP_REPAIR_SKIPPED", "message": str(exc)}]
