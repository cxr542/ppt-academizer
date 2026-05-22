"""Slide limit / quality mode tests."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.slide_limits import (  # noqa: E402
    slide_limit_payload,
    standard_max_slides,
    validate_quality_mode,
)


def test_standard_blocks_over_limit() -> None:
    max_std = standard_max_slides()
    try:
        validate_quality_mode("standard", source_slides=max_std + 1, estimated_output=None)
        assert False, "expected ValueError"
    except ValueError as e:
        assert "표준" in str(e)


def test_unlimited_allows_over_limit() -> None:
    max_std = standard_max_slides()
    mode = validate_quality_mode(
        "unlimited", source_slides=max_std + 50, estimated_output=max_std + 100
    )
    assert mode == "unlimited"


def test_payload_flags_over_limit() -> None:
    p = slide_limit_payload(source_slides=200, estimated_output=198)
    assert p["over_standard_limit"] is True
    assert p["default_quality_mode"] == "unlimited"
    assert p["quality_modes"][0]["available"] is False
    assert p["quality_modes"][1]["available"] is True
