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


def test_save_upload_streams_to_disk_and_keeps_the_bytes(monkeypatch: pytest.MonkeyPatch) -> None:
    from api.upload_handlers import remove_upload_tmp, save_upload_bounded

    monkeypatch.setattr("api.upload_handlers.max_upload_bytes", lambda: 10_000)
    monkeypatch.setattr("api.upload_handlers._READ_CHUNK", 16)  # force several chunks
    payload = b"PK" + bytes(range(200))
    upload = UploadFile(filename="deck.pptx", file=io.BytesIO(payload))

    tmp, src = asyncio.run(save_upload_bounded(upload))
    try:
        assert src.name == "deck.pptx" and src.read_bytes() == payload
    finally:
        remove_upload_tmp(tmp)


def test_save_upload_rejects_oversized_file_and_removes_the_temp_dir(monkeypatch: pytest.MonkeyPatch) -> None:
    import tempfile
    from pathlib import Path

    from api.upload_handlers import save_upload_bounded

    monkeypatch.setattr("api.upload_handlers.max_upload_bytes", lambda: 100)
    monkeypatch.setattr("api.upload_handlers._READ_CHUNK", 32)
    created: list[Path] = []
    real_mkdtemp = tempfile.mkdtemp

    def spy(*a, **k):
        path = real_mkdtemp(*a, **k)
        created.append(Path(path))
        return path

    monkeypatch.setattr("api.upload_handlers.tempfile.mkdtemp", spy)
    upload = UploadFile(filename="big.pptx", file=io.BytesIO(b"x" * 250))

    with pytest.raises(HTTPException) as exc:
        asyncio.run(save_upload_bounded(upload))
    assert exc.value.status_code == 413
    assert created and not created[0].exists()
