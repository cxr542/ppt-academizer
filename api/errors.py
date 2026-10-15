"""API error helpers."""

from __future__ import annotations

import logging
import os
import traceback
import uuid

from zipfile import BadZipFile

from fastapi import HTTPException
from pptx.exc import PackageNotFoundError

_log = logging.getLogger("ppt-academizer.api")


def _debug_errors() -> bool:
    return os.environ.get("PPT_ACADEMIZER_DEBUG_ERRORS", "").strip().lower() in ("1", "true", "yes", "on")


def internal_server_error(exc: Exception) -> HTTPException:
    """500 for an unexpected failure.

    The client only gets a short Korean message and an error code; the full traceback (with server
    paths and library internals) stays in the server log under the same code. Set
    ``PPT_ACADEMIZER_DEBUG_ERRORS=1`` locally to get the traceback in the response too.
    """
    code = uuid.uuid4().hex[:8]
    _log.exception("request failed [%s]: %s", code, exc)
    detail = f"처리 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요. (오류 코드: {code})"
    if _debug_errors():
        detail += "\n\n" + "".join(traceback.format_exception(exc))
    return HTTPException(status_code=500, detail=detail)


def missing_resource_error(exc: Exception, detail: str, status_code: int = 400) -> HTTPException:
    """A missing file/template: tell the client what to do, keep the server path in the log only."""
    _log.warning("missing resource: %s", exc)
    return HTTPException(status_code=status_code, detail=detail)


# python-pptx raises these for a file that is not a readable .pptx package.
UNREADABLE_PPTX_ERRORS = (BadZipFile, PackageNotFoundError)


def unreadable_pptx_error(exc: Exception) -> HTTPException:
    """The upload is not a valid .pptx (corrupt, truncated, or another format renamed)."""
    _log.warning("unreadable pptx: %s", exc)
    return HTTPException(
        status_code=400,
        detail="PowerPoint(.pptx) 파일을 읽을 수 없습니다. 파일이 손상되었거나 .pptx 형식이 아닐 수 있습니다. "
        "PowerPoint에서 열어 다시 저장한 뒤 올려 주세요.",
    )
