"""Warning formatting for UI."""

import re

from core.warnings_display import format_warning, format_warnings


def test_korean_pipeline_message() -> None:
    w = format_warning({"code": "PIPELINE_MIGRATE_CMP", "message": ""})
    assert "§7" in w["message"] or "도형" in w["message"]
    assert w["level"] == "info"


def test_slide_kept_empty_uses_1_based_slide() -> None:
    w = format_warning({"code": "SLIDE_KEPT_EMPTY", "src_index": 22})
    assert "23" in w["message"]


def test_traceback_hidden() -> None:
    w = format_warning(
        {
            "code": "OOXML_VALIDATE_FAILED",
            "message": "Traceback (most recent call last):\n  File validate.py",
        }
    )
    assert "Traceback" not in w["message"]
    assert "PowerPoint" in w["message"] or "검증" in w["message"]


def test_syntax_error_hidden() -> None:
    w = format_warning(
        {
            "code": "OOXML_VALIDATE_FAILED",
            "message": (
                'File "engine/office/validate.py", line 80\n'
                "    match file_extension:\n"
                "SyntaxError: invalid syntax"
            ),
        }
    )
    assert "SyntaxError" not in w["message"]
    assert "match file_extension" not in w["message"]
    assert "검증" in w["message"] or "PowerPoint" in w["message"]


def test_validate_unavailable_token_hidden() -> None:
    w = format_warning(
        {
            "code": "OOXML_VALIDATE_FAILED",
            "message": "validate_pptx_unavailable",
        }
    )
    assert "validate_pptx_unavailable" not in w["message"]
    assert "PowerPoint" in w["message"] or "검증" in w["message"]


def test_repair_import_error_hidden() -> None:
    w = format_warning(
        {
            "code": "OOXML_REPAIR_FAILED",
            "message": "dlopen(pyexpat): Symbol not found",
        }
    )
    assert "dlopen" not in w["message"]
    assert "자동 정리" in w["message"]
    assert w["level"] == "warn"


def test_format_warnings_skips_meta() -> None:
    out = format_warnings([{"code": "MIGRATE_META", "deck_kind": "cmp"}])
    assert out == []


_HANGUL = re.compile(r"[가-힣]")


def test_engine_warning_codes_never_leak_english_internal_text() -> None:
    # Codes the engine emits with English/internal messages (paths, counts) must be rewritten.
    engine_msgs = {
        "BRAND_PACK": "Using academy brand pack at /tmp/brand",
        "BRAND_PACK_MISSING": "brand/ pack not found next to template; using built-in theme colors.",
        "CANVAS_RESCALED": "Source slide size 1×2 differs from academy 3×4; body shapes were scaled to fit.",
        "LAYOUT_PATTERN": "Applied academy 3col card pattern on content step 2.",
        "SOURCE_PLACEHOLDER_IMAGE": "Source slide 3: 1 placeholder image(s) (e.g. 600×400 stub)",
        "BRAND_ICONS_REPLACED": "Replaced 2 icon(s) with academy brand pack: a, b",
    }
    for code, msg in engine_msgs.items():
        out = format_warning({"code": code, "message": msg, "count": 2})
        assert _HANGUL.search(out["message"]), code
        assert "/tmp" not in out["message"], code
