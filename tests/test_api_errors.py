"""Unexpected server errors must not leak tracebacks, paths or internals to the client."""

import logging

from api.errors import internal_server_error


def _boom():
    raise RuntimeError("secret internal detail at /app/engine/scripts/x.py")


def _raise_and_wrap():
    try:
        _boom()
    except RuntimeError as exc:
        return internal_server_error(exc)


def test_client_message_is_short_korean_with_a_code(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_DEBUG_ERRORS", raising=False)
    err = _raise_and_wrap()
    assert err.status_code == 500
    assert "처리 중 오류가 발생했습니다" in err.detail
    assert "오류 코드:" in err.detail
    for leaked in ("Traceback", "/app/engine", "secret internal detail", "RuntimeError", ".py"):
        assert leaked not in err.detail, leaked


def test_the_same_code_is_in_the_server_log_with_the_traceback(monkeypatch, caplog) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_DEBUG_ERRORS", raising=False)
    with caplog.at_level(logging.ERROR, logger="ppt-academizer.api"):
        err = _raise_and_wrap()
    code = err.detail.split("오류 코드: ")[1].rstrip(")")
    record = next(r for r in caplog.records if code in r.getMessage())
    assert record.exc_info is not None and "secret internal detail" in record.getMessage()


def test_debug_switch_adds_the_traceback(monkeypatch) -> None:
    monkeypatch.setenv("PPT_ACADEMIZER_DEBUG_ERRORS", "1")
    err = _raise_and_wrap()
    assert "Traceback" in err.detail and "secret internal detail" in err.detail


def test_codes_differ_between_errors(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_DEBUG_ERRORS", raising=False)
    assert _raise_and_wrap().detail != _raise_and_wrap().detail


def test_missing_resource_keeps_the_path_out_of_the_response(caplog) -> None:
    import logging

    from api.errors import missing_resource_error

    exc = FileNotFoundError("/app/engine/templates/academy-template.pptx")
    with caplog.at_level(logging.WARNING, logger="ppt-academizer.api"):
        err = missing_resource_error(exc, "서버에서 아카데미 템플릿을 찾을 수 없습니다.", 503)
    assert err.status_code == 503 and "/app" not in err.detail
    assert any("/app/engine/templates" in r.getMessage() for r in caplog.records)


def test_unreadable_pptx_is_a_400_with_guidance() -> None:
    from zipfile import BadZipFile

    from pptx.exc import PackageNotFoundError

    from api.errors import UNREADABLE_PPTX_ERRORS, unreadable_pptx_error

    assert BadZipFile in UNREADABLE_PPTX_ERRORS and PackageNotFoundError in UNREADABLE_PPTX_ERRORS
    err = unreadable_pptx_error(BadZipFile("File is not a zip file"))
    assert err.status_code == 400
    assert ".pptx" in err.detail and "zip" not in err.detail.lower()
