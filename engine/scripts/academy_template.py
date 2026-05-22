"""Resolve path to OKESTRO academy 2026 lecture template .pptx."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path

DEFAULT_TEMPLATE_BASENAME = "1.아카데미 강의안 템플릿(2026).pptx"


def resolve_academy_template_path() -> Path:
    """Return absolute path to the academy template file.

    Resolution order:
    1. Environment variable ``TEMPLATE_PPTX`` (absolute or relative path).
    2. macOS Spotlight ``mdfind`` by default basename.

    Raises:
        FileNotFoundError: If no file is found or path is not a file.
    """
    env = os.environ.get("TEMPLATE_PPTX", "").strip()
    if env:
        p = Path(env).expanduser().resolve()
        if not p.is_file():
            raise FileNotFoundError(
                f"TEMPLATE_PPTX is set but not a file: {p}\n"
                "Unset TEMPLATE_PPTX or point it to the academy .pptx file."
            )
        return p

    r = subprocess.run(
        ["mdfind", "-name", DEFAULT_TEMPLATE_BASENAME],
        capture_output=True,
        text=True,
        check=False,
    )
    if r.returncode != 0 or not r.stdout.strip():
        raise FileNotFoundError(
            "템플릿을 찾지 못했습니다. 다음 중 하나를 하세요:\n"
            "  • OneDrive 동기화 후 다시 실행 (Spotlight 인덱싱)\n"
            "  • 템플릿을 docs/_academy_template_2026.pptx 등으로 복사 후\n"
            f"    export TEMPLATE_PPTX=\"/절대경로/{DEFAULT_TEMPLATE_BASENAME}\""
        )
    p = Path(r.stdout.strip().split("\n")[0])
    if not p.is_file():
        raise FileNotFoundError(p)
    return p
