"""User-facing warning text for web UI (Korean, no tracebacks)."""

from __future__ import annotations

import re

_TRACEBACK_RE = re.compile(r"Traceback \(most recent call last\):", re.I)


def _src_slide_label(w: dict) -> int | None:
    """1-based source slide number for users."""
    if w.get("slide") is not None:
        return int(w["slide"])
    if w.get("src_index") is not None:
        return int(w["src_index"]) + 1
    return None


def _sanitize_technical_message(msg: str, code: str) -> str:
    if not msg:
        return ""
    if (
        _TRACEBACK_RE.search(msg)
        or "ModuleNotFoundError" in msg
        or "ImportError" in msg
        or "SyntaxError" in msg
    ):
        if code == "OOXML_VALIDATE_FAILED":
            return (
                "PPTX 구조 자동 검증을 실행하지 못했습니다. "
                "변환 파일은 저장되었으니 PowerPoint에서 열어 확인해 주세요."
            )
        return "내부 검증 단계에서 오류가 있었습니다. 변환 파일은 저장되었습니다."
    if len(msg) > 240:
        return msg[:237] + "…"
    return msg


def format_warning(w: dict) -> dict:
    """Return warning dict with `message` (Korean) and optional `level`."""
    code = str(w.get("code") or "MIGRATE")
    slide_n = _src_slide_label(w)
    prefix = f"원본 {slide_n}장: " if slide_n is not None else ""

    templates: dict[str, str | tuple[str, str]] = {
        "PIPELINE_MIGRATE_CMP": ("info", "도형 이식(§7) 방식으로 변환했습니다."),
        "ACADEMY_SOURCE_LAYOUT": (
            "info",
            "아카데미 마스터로 만든 원본입니다. 헤더·도형을 §7 규칙으로 다시 맞춥니다.",
        ),
        "QUALITY_MODE_UNLIMITED": (
            "warn",
            "대용량 모드: 슬라이드가 많아 레이아웃·제목·도형 품질을 보장하지 않습니다. 결과는 반드시 검수하세요.",
        ),
        "SLIDE_KEPT_EMPTY": (
            "info",
            "내용이 거의 없어 빈 본문 슬라이드로 넣었습니다. 이미지·차트는 §6.7 규칙으로 옮깁니다.",
        ),
        "CHART_NOT_COPIED": (
            "warn",
            "차트·그래프 일부를 옮기지 못했습니다. 원본을 이미지로 붙여 넣은 경우도 있습니다.",
        ),
        "RASTER_DIAGRAM_RESTORED": (
            "info",
            "본문 이미지·다이어그램을 추가로 옮겼습니다.",
        ),
        "SLIDE_COUNT_MISMATCH": (
            "warn",
            "출력 슬라이드 수가 예상과 다릅니다. 목차 삽입 여부를 확인해 주세요.",
        ),
        "OOXML_VALIDATE_FAILED": (
            "info",
            _sanitize_technical_message(str(w.get("message") or ""), code)
            or "PPTX 구조 자동 검증에 실패했습니다. 파일은 저장되었습니다.",
        ),
        "OOXML_VALIDATE_SKIPPED": ("info", "PPTX 구조 자동 검증은 건너뛰었습니다."),
        "SVG_RASTERIZED": ("info", ""),
        "PP_REPAIR_APPLIED": (
            "info",
            "Mac PowerPoint 호환 정리를 적용했습니다(이미지·SVG 패키지 정리).",
        ),
        "PP_REPAIR_SKIPPED": (
            "warn",
            "Mac PowerPoint 자동 복구를 적용하지 못했습니다. "
            "열 때 복구 창이 뜨면「복구」를 누른 뒤 저장하거나, "
            "Microsoft PowerPoint가 설치·실행 가능한지 확인해 주세요. "
            "(SVG·파트너 덱에서 자주 발생)",
        ),
        "TEXT_NOT_EXTRACTED": ("warn", ""),
    }

    level = "info"
    body = _sanitize_technical_message(str(w.get("message") or ""), code)

    if code in templates:
        entry = templates[code]
        if isinstance(entry, tuple):
            level, tpl_body = entry
            if tpl_body:
                body = tpl_body
        else:
            body = entry

    if code not in templates and body:
        body = _sanitize_technical_message(body, code)

    return {
        "code": code,
        "level": level,
        "slide": slide_n,
        "message": prefix + body if prefix and not body.startswith("원본") else body,
    }


def format_warnings(warnings: list[dict]) -> list[dict]:
    """Drop internal meta; format for UI."""
    out: list[dict] = []
    for w in warnings:
        if w.get("code") in ("MIGRATE_META",):
            continue
        out.append(format_warning(w))
    return out
