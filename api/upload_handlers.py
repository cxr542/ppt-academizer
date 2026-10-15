"""Shared upload helpers for API routes (.pptx only)."""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

from fastapi import HTTPException, UploadFile

from core.upload_limits import max_upload_bytes

_READ_CHUNK = 1024 * 1024


def require_pptx_filename(filename: str) -> None:
    name = (filename or "").lower()
    if not name.endswith(".pptx"):
        raise HTTPException(
            status_code=400,
            detail="입력은 PowerPoint(.pptx) 파일만 지원합니다.",
        )


async def read_upload_bounded(file: UploadFile) -> tuple[bytes, str]:
    if not file.filename:
        raise HTTPException(status_code=400, detail="파일 이름이 없습니다.")
    require_pptx_filename(file.filename)
    limit = max_upload_bytes()
    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = await file.read(_READ_CHUNK)
        if not chunk:
            break
        total += len(chunk)
        if total > limit:
            from core.upload_limits import file_too_large_message

            raise HTTPException(status_code=413, detail=file_too_large_message())
        chunks.append(chunk)
    return b"".join(chunks), file.filename


async def save_upload_bounded(file: UploadFile) -> tuple[Path, Path]:
    """Stream the upload straight to a temp file (never holds the whole deck in memory).

    Returns ``(tmp_dir, source_path)``; the temp dir is removed again if the upload is rejected.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="파일 이름이 없습니다.")
    require_pptx_filename(file.filename)
    limit = max_upload_bytes()
    tmp = Path(tempfile.mkdtemp(prefix="ppt-academizer-upload-"))
    src = tmp / Path(file.filename).name
    total = 0
    try:
        with src.open("wb") as out:
            while True:
                chunk = await file.read(_READ_CHUNK)
                if not chunk:
                    break
                total += len(chunk)
                if total > limit:
                    from core.upload_limits import file_too_large_message

                    raise HTTPException(status_code=413, detail=file_too_large_message())
                out.write(chunk)
    except BaseException:
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    return tmp, src


def write_upload_tmp(content: bytes, filename: str) -> tuple[Path, Path]:
    require_pptx_filename(filename)
    tmp = Path(tempfile.mkdtemp(prefix="ppt-academizer-upload-"))
    src = tmp / Path(filename).name
    src.write_bytes(content)
    return tmp, src


def remove_upload_tmp(tmp: Path | None) -> None:
    if tmp is None:
        return
    shutil.rmtree(tmp, ignore_errors=True)
