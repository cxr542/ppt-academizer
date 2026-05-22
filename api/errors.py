"""API error helpers."""

from __future__ import annotations

import logging

from fastapi import HTTPException

_log = logging.getLogger("ppt-academizer.api")


def internal_server_error(exc: Exception) -> HTTPException:
    _log.exception("request failed: %s", exc)
    return HTTPException(status_code=500, detail="처리 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.")
