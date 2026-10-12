"""Unit tests for deck_kind migrate configuration."""

from __future__ import annotations

from scripts.deck_migrate_config import config_for_kind, detect_deck_kind


def test_detect_deck_kind_cmp() -> None:
    assert detect_deck_kind("클라우드 구현기술(CMP)_v1.pptx") == "cmp"


def test_detect_deck_kind_contrabass() -> None:
    assert detect_deck_kind("CONTRABASS 기반기술@260504.pptx") == "contrabass"


def test_cmp_auto_toc() -> None:
    cfg = config_for_kind("cmp")
    assert cfg.auto_toc_after_cover is True
    assert 8 in cfg.skip_src_indices


def test_contrabass_no_cmp_toc() -> None:
    cfg = config_for_kind("contrabass")
    assert cfg.auto_toc_after_cover is False
    assert cfg.skip_src_indices == frozenset()
    assert cfg.cover_subtitle is None
