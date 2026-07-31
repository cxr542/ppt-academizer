"""Warning formatting for UI."""

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


def test_format_warnings_skips_meta() -> None:
    out = format_warnings([{"code": "MIGRATE_META", "deck_kind": "cmp"}])
    assert out == []
