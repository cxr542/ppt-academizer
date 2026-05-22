"""Smoke tests for wizard preview payload."""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def test_build_wizard_preview_structure():
    from core.wizard_preview import build_wizard_preview

    fake_src = Path("/tmp/fake-deck.pptx")
    spec_preview = {
        "output_slide_count": 10,
        "summary": ["출력 예상 10장"],
        "slides": [{"out_slide": 1, "layout": "2_표지", "lines": ["제목"]}],
    }
    migrate_preview = {
        "output_slide_count": 8,
        "summary": ["출력 예상 약 8장"],
        "slides": [{"src_slide": 1, "plan_kind": "cover", "layout_hint": "2_표지", "lines": ["표지"]}],
    }

    with (
        patch("core.wizard_preview.Presentation") as mock_prs,
        patch("core.wizard_preview.detect_deck_profile", return_value=("migrate_cmp", {"reason": "shape-heavy"})),
        patch("core.wizard_preview.analyze_deck_structure") as mock_st,
        patch("core.wizard_preview._preview_spec", return_value=spec_preview),
        patch("core.wizard_preview._preview_migrate", return_value=migrate_preview),
    ):
        mock_st.return_value = MagicMock(
            slide_count=20,
            shapes_per_slide=7.5,
            picture_count=5,
            table_count=2,
            chart_count=1,
            total_text_blocks=40,
            google_image_ratio=0.1,
        )
        mock_prs.return_value = MagicMock()
        out = build_wizard_preview(fake_src)

    assert out["detected_profile"] == "migrate_cmp"
    assert out["recommended_profile"] == "migrate_cmp"
    assert len(out["cards"]) == 2
    assert out["cards"][0]["id"] == "migrate_cmp"
    assert out["cards"][0]["recommended"] is True
    assert out["cards"][0]["preview"] == migrate_preview
    assert out["cards"][1]["preview"] == spec_preview
    assert len(out["structure_bullets"]) >= 2
