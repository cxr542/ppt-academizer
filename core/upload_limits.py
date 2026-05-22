"""Upload size limits for API endpoints."""

from __future__ import annotations

import os

# Stabilization defaults (expand via env after quality is proven)
_DEFAULT_MAX_MB = 50


def max_upload_bytes() -> int:
    raw = os.environ.get("PPT_ACADEMIZER_MAX_UPLOAD_MB", str(_DEFAULT_MAX_MB)).strip()
    try:
        mb = int(raw)
    except ValueError:
        mb = _DEFAULT_MAX_MB
    mb = max(5, min(mb, 150))
    return mb * 1024 * 1024


def max_upload_mb() -> int:
    return max_upload_bytes() // (1024 * 1024)


def file_too_large_message() -> str:
    return f"파일 크기는 {max_upload_mb()}MB 이하여야 합니다."
