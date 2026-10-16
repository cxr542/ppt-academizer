#!/usr/bin/env python3
"""Per-deck migrate settings (§7).

Built-in kinds below are defaults. Partner-specific values (skip slides, part covers, cover
subtitle, filename patterns) can be added or overridden without code changes through a JSON file:

  1. ``$PPT_ACADEMIZER_DECK_KINDS`` (explicit path), else
  2. ``<app root>/config/deck_kinds.json`` (outside ``engine/``, so engine syncs keep it).

See ``config/deck_kinds.example.json``. A malformed file is ignored (built-ins are used).
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

DeckKind = str  # built-ins: cmp | contrabass | paas | generic_partner; JSON may add more

GENERIC_KIND = "generic_partner"
_AI_PATTERN = r"chatgpt|openai|gemini|copilot|claude|gamma\.?app|ai[-_ ]?(deck|ppt|export)"


@dataclass(frozen=True)
class DeckMigrateConfig:
    deck_kind: DeckKind
    skip_src_indices: frozenset[int]
    auto_toc_after_cover: bool
    toc_slide_range: range | None
    cover_subtitle: str | None
    part_cover_indices: frozenset[int]


# Detection order matters: AI exports first (their filenames may also contain other words).
_BUILTIN_KINDS: dict[str, dict[str, Any]] = {
    GENERIC_KIND: {
        "filename_patterns": [_AI_PATTERN],
        "skip_src_indices": [],
        "auto_toc_after_cover": False,
        "toc_slide_range": None,
        "cover_subtitle": None,
        "part_cover_indices": [0],
    },
    "contrabass": {
        "filename_patterns": [r"contrabass|콘트라베이스"],
        "skip_src_indices": [],
        "auto_toc_after_cover": False,
        "toc_slide_range": None,
        "cover_subtitle": None,
        "part_cover_indices": [0],
    },
    "cmp": {
        "filename_patterns": [r"\bcmp\b|클라우드\s*구현"],
        "skip_src_indices": [8],
        "auto_toc_after_cover": True,
        "toc_slide_range": [2, 8],
        "cover_subtitle": "클라우드 구현기술 (CMP)",
        "part_cover_indices": [0, 13],
    },
    "paas": {
        "filename_patterns": [r"\bpaas\b"],
        "skip_src_indices": [],
        "auto_toc_after_cover": False,
        "toc_slide_range": None,
        "cover_subtitle": None,
        "part_cover_indices": [0],
    },
}

# Shape names (Google exports) that are never body content. Extend via JSON.
_BUILTIN_SKIP_SHAPE_NAMES = frozenset({"Google Shape;520;p15"})

# Titles that mark the last slide as a closing ("thank you") slide. Extend/replace via the JSON
# key ``closing_slide_patterns`` (list of regexes, case-insensitive); ``[]`` turns the rule off.
_BUILTIN_CLOSING_PATTERNS = (
    r"감사합니다|감사드립니다|고맙습니다",
    r"thank\s*you|thanks\s+for",
    r"^\s*q\s*&\s*a\s*$",
)

# The template's 간지/목차 layouts print the author's example course name in a fixed text box. A
# converted deck is about another course, so that text is replaced with the deck's own title.
# Override with the JSON key ``template_course_name_text``; an empty string leaves the layouts as is.
_BUILTIN_TEMPLATE_COURSE_NAME = "Introduction to CLOUD COMPUTING"

_APP_ROOT = Path(__file__).resolve().parents[2]


def _config_path() -> Path | None:
    env = os.environ.get("PPT_ACADEMIZER_DECK_KINDS", "").strip()
    if env:
        return Path(env).expanduser()
    default = _APP_ROOT / "config" / "deck_kinds.json"
    return default if default.is_file() else None


@lru_cache(maxsize=8)
def _load_overrides(path: str | None, mtime: float | None) -> dict[str, Any]:
    if not path:
        return {}
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def _overrides() -> dict[str, Any]:
    path = _config_path()
    if path is None or not path.is_file():
        return {}
    return _load_overrides(str(path), path.stat().st_mtime)


def _kind_table() -> dict[str, dict[str, Any]]:
    """User kinds first (so they win filename matching), then built-ins with overrides applied."""
    user = _overrides().get("kinds")
    user = user if isinstance(user, dict) else {}
    table: dict[str, dict[str, Any]] = {}
    for name, spec in user.items():
        if name not in _BUILTIN_KINDS and isinstance(spec, dict):
            table[name] = {**_BUILTIN_KINDS[GENERIC_KIND], "filename_patterns": [], **spec}
    for name, base in _BUILTIN_KINDS.items():
        override = user.get(name)
        table[name] = {**base, **override} if isinstance(override, dict) else dict(base)
    return table


def skip_shape_names() -> frozenset[str]:
    extra = _overrides().get("global_skip_shape_names")
    if isinstance(extra, list):
        return _BUILTIN_SKIP_SHAPE_NAMES | frozenset(str(x) for x in extra)
    return _BUILTIN_SKIP_SHAPE_NAMES


def template_course_name_text() -> str:
    configured = _overrides().get("template_course_name_text")
    return configured.strip() if isinstance(configured, str) else _BUILTIN_TEMPLATE_COURSE_NAME


def footer_copyright_year(today_year: int | None = None) -> int | None:
    """Year printed in the layout footer ("© 2026. OKESTRO …").

    JSON key ``footer_copyright_year``: ``"auto"`` (default) = the year of the conversion; an
    integer pins a year; ``null`` or ``""`` leaves the template's footer untouched.
    """
    import datetime

    configured = _overrides().get("footer_copyright_year", "auto")
    if configured is None or configured == "":
        return None
    if isinstance(configured, bool):
        return today_year or datetime.date.today().year
    if isinstance(configured, int):
        return configured
    if isinstance(configured, str) and configured.strip().isdigit():
        return int(configured.strip())
    return today_year or datetime.date.today().year


def closing_slide_patterns() -> tuple[re.Pattern[str], ...]:
    configured = _overrides().get("closing_slide_patterns")
    raw = configured if isinstance(configured, list) else _BUILTIN_CLOSING_PATTERNS
    compiled: list[re.Pattern[str]] = []
    for pattern in raw:
        try:
            compiled.append(re.compile(str(pattern), re.I))
        except re.error:
            continue
    return tuple(compiled)


def detect_deck_kind(filename: str) -> DeckKind:
    name = filename.lower()
    for kind, spec in _kind_table().items():
        for pattern in spec.get("filename_patterns") or []:
            try:
                if re.search(pattern, name, re.I):
                    return kind
            except re.error:
                continue
    return GENERIC_KIND


def _to_config(kind: str, spec: dict[str, Any]) -> DeckMigrateConfig:
    toc = spec.get("toc_slide_range")
    return DeckMigrateConfig(
        deck_kind=kind,
        skip_src_indices=frozenset(int(i) for i in spec.get("skip_src_indices") or []),
        auto_toc_after_cover=bool(spec.get("auto_toc_after_cover", False)),
        toc_slide_range=range(int(toc[0]), int(toc[1])) if toc else None,
        cover_subtitle=spec.get("cover_subtitle") or None,
        # An empty set would silently lose the cover; slide 0 is always a cover.
        part_cover_indices=frozenset(int(i) for i in spec.get("part_cover_indices") or []) | {0},
    )


def config_for_kind(kind: DeckKind) -> DeckMigrateConfig:
    table = _kind_table()
    return _to_config(kind if kind in table else GENERIC_KIND, table.get(kind, table[GENERIC_KIND]))


def migrate_config_for_source(source: Path | str) -> DeckMigrateConfig:
    path = Path(source)
    return config_for_kind(detect_deck_kind(path.name))
