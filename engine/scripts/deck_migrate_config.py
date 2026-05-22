#!/usr/bin/env python3
"""Per-deck migrate settings (§7) — replaces CMP-only hardcoding in build_cmp_academy."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

DeckKind = str  # cmp | contrabass | paas | generic_partner


@dataclass(frozen=True)
class DeckMigrateConfig:
    deck_kind: DeckKind
    skip_src_indices: frozenset[int]
    auto_toc_after_cover: bool
    toc_slide_range: range | None
    cover_subtitle: str | None
    part_cover_indices: frozenset[int]


def detect_deck_kind(filename: str) -> DeckKind:
    name = filename.lower()
    if re.search(
        r"chatgpt|openai|gemini|copilot|claude|gamma\.?app|ai[-_ ]?(deck|ppt|export)",
        name,
        re.I,
    ):
        return "generic_partner"
    if re.search(r"contrabass|콘트라베이스", name, re.I):
        return "contrabass"
    if re.search(r"\bcmp\b|클라우드\s*구현", name, re.I):
        return "cmp"
    if re.search(r"\bpaas\b", name, re.I):
        return "paas"
    return "generic_partner"


def config_for_kind(kind: DeckKind) -> DeckMigrateConfig:
    if kind == "cmp":
        return DeckMigrateConfig(
            deck_kind="cmp",
            skip_src_indices=frozenset({8}),
            auto_toc_after_cover=True,
            toc_slide_range=range(2, 8),
            cover_subtitle="클라우드 구현기술 (CMP)",
            part_cover_indices=frozenset({0, 13}),
        )
    if kind == "contrabass":
        return DeckMigrateConfig(
            deck_kind="contrabass",
            skip_src_indices=frozenset(),
            auto_toc_after_cover=False,
            toc_slide_range=None,
            cover_subtitle=None,
            part_cover_indices=frozenset({0}),
        )
    if kind == "paas":
        return DeckMigrateConfig(
            deck_kind="paas",
            skip_src_indices=frozenset(),
            auto_toc_after_cover=False,
            toc_slide_range=None,
            cover_subtitle=None,
            part_cover_indices=frozenset({0}),
        )
    return DeckMigrateConfig(
        deck_kind="generic_partner",
        skip_src_indices=frozenset(),
        auto_toc_after_cover=False,
        toc_slide_range=None,
        cover_subtitle=None,
        part_cover_indices=frozenset({0}),
    )


def migrate_config_for_source(source: Path | str) -> DeckMigrateConfig:
    path = Path(source)
    return config_for_kind(detect_deck_kind(path.name))
