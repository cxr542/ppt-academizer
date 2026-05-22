"""Analyze source deck and spec path before academize."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from .ppt_test_path import ensure_engine_on_path

ensure_engine_on_path()

from pptx import Presentation  # noqa: E402

from scripts.convert_legacy_deck_to_academy import (  # noqa: E402
    analyze_specs,
    convert_presentation,
)
from scripts.deck_profile import detect_deck_profile  # noqa: E402


def analyze_presentation(
    source: Path,
    *,
    deck_title: str | None = None,
    deck_subtitle: str = "PPT 아카데미화",
    profile: str = "auto",
) -> dict[str, Any]:
    """Return profile, routing hint, and per-slide spec preview (spec path)."""
    source = Path(source).resolve()
    if not source.is_file():
        raise FileNotFoundError(source)

    title = deck_title or source.stem
    prs = Presentation(str(source))
    detected, meta = detect_deck_profile(prs, filename_hint=source.name)

    if profile == "auto":
        route = detected
    else:
        route = profile

    specs = convert_presentation(
        source,
        deck_title=title,
        deck_subtitle=deck_subtitle,
        include_default_front_matter=(route == "spec" or route == "google_image"),
        front_matter_mode="auto",
    )
    slides = analyze_specs(specs)
    validation_errors = [issue for row in slides for issue in row.get("validation", [])]

    return {
        "source": str(source),
        "deck_title": title,
        "detected_profile": detected,
        "profile_meta": meta,
        "route_profile": route,
        "recommended_pipeline": (
            "migrate_cmp (academy-design §7)"
            if route == "migrate_cmp"
            else "spec_json (academy-design §5)"
        ),
        "spec_slide_count": len(specs),
        "slides": slides,
        "validation_errors": validation_errors,
    }
