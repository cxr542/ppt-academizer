"""Upload size enforcement tests."""

from __future__ import annotations

import asyncio
import io

import pytest
from fastapi import HTTPException
from starlette.datastructures import UploadFile

from api.upload_handlers import read_upload_bounded


def test_read_upload_rejects_oversized_file(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("api.upload_handlers.max_upload_bytes", lambda: 100)
    upload = UploadFile(filename="big.pptx", file=io.BytesIO(b"x" * 250))

    async def _run() -> None:
        await read_upload_bounded(upload)

    with pytest.raises(HTTPException) as exc:
        asyncio.run(_run())
    assert exc.value.status_code == 413


def test_read_upload_accepts_within_limit(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("api.upload_handlers.max_upload_bytes", lambda: 100)
    payload = b"PK" + b"\x00" * 50
    upload = UploadFile(filename="ok.pptx", file=io.BytesIO(payload))

    async def _run() -> tuple[bytes, str]:
        return await read_upload_bounded(upload)

    content, name = asyncio.run(_run())
    assert name == "ok.pptx"
    assert len(content) == len(payload)
