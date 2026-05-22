"""Slide-count tiers: standard (quality) vs unlimited (best-effort)."""

from __future__ import annotations

import os
from typing import Any, Literal

QualityMode = Literal["standard", "unlimited"]

# Stabilization: quality-first cap (raise via PPT_ACADEMIZER_MAX_SLIDES_STANDARD later)
_DEFAULT_STANDARD_MAX = 40


def standard_max_slides() -> int:
    raw = os.environ.get("PPT_ACADEMIZER_MAX_SLIDES_STANDARD", str(_DEFAULT_STANDARD_MAX)).strip()
    try:
        n = int(raw)
    except ValueError:
        n = _DEFAULT_STANDARD_MAX
    return max(5, min(n, 200))


def effective_slide_count(*, source_slides: int, estimated_output: int | None = None) -> int:
    """Conservative budget for limit checks (migrate can match or exceed source)."""
    if estimated_output is not None and estimated_output > 0:
        return max(source_slides, estimated_output)
    return source_slides


def slide_limit_payload(
    *,
    source_slides: int,
    estimated_output: int | None = None,
) -> dict[str, Any]:
    max_std = standard_max_slides()
    budget = effective_slide_count(
        source_slides=source_slides, estimated_output=estimated_output
    )
    over = budget > max_std
    return {
        "source_slide_count": source_slides,
        "estimated_output_slides": estimated_output,
        "effective_slide_count": budget,
        "standard_max_slides": max_std,
        "over_standard_limit": over,
        "quality_modes": [
            {
                "id": "standard",
                "title": "표준 (안정·품질 권장)",
                "description": (
                    f"원본·예상 출력 {max_std}장 이하. 변환 품질을 우선합니다. "
                    "긴 덱은 파트별로 나누는 것을 권장합니다."
                ),
                "available": not over,
            },
            {
                "id": "unlimited",
                "title": "대용량 (실험·품질 미보장)",
                "description": "장수 제한 없이 변환합니다. 안정화 단계에서는 비권장입니다.",
                "disclaimer": (
                    "슬라이드가 많을수록 레이아웃·제목·도형 누락이 늘어납니다. "
                    f"{max_std}장 이하로 나눈 뒤 표준 모드를 사용해 주세요."
                ),
                "available": True,
            },
        ],
        "default_quality_mode": "unlimited" if over else "standard",
    }


def validate_quality_mode(
    mode: str,
    *,
    source_slides: int,
    estimated_output: int | None = None,
) -> QualityMode:
    if mode not in ("standard", "unlimited"):
        raise ValueError("quality_mode는 standard 또는 unlimited 여야 합니다.")
    q: QualityMode = mode  # type: ignore[assignment]
    if q == "standard":
        max_std = standard_max_slides()
        budget = effective_slide_count(
            source_slides=source_slides, estimated_output=estimated_output
        )
        if budget > max_std:
            raise ValueError(
                f"이 덱은 약 {budget}장 기준으로 표준 모드 한도({max_std}장)를 초과합니다. "
                f"「대용량(품질 미보장)」을 선택하거나, 파트별로 파일을 나눠 주세요."
            )
    return q
