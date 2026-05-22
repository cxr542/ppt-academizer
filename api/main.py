#!/usr/bin/env python3
"""ppt-academizer — minimal web UI + /academize API (.pptx converter)."""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import os
import sys
from pathlib import Path
from typing import Callable, TypeVar

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from api.errors import internal_server_error  # noqa: E402
from api.upload_handlers import (  # noqa: E402
    read_upload_bounded,
    remove_upload_tmp,
    write_upload_tmp,
)
from core.ppt_test_path import ensure_engine_on_path  # noqa: E402
from core.slide_limits import standard_max_slides  # noqa: E402
from core.upload_limits import max_upload_bytes, max_upload_mb  # noqa: E402
from core.version import LATEST_RELEASE_DOC, SERVICE_VERSION  # noqa: E402

ENGINE_ROOT = ensure_engine_on_path()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

logging.basicConfig(level=logging.INFO)

WEB = ROOT / "web"
app = FastAPI(title="ppt-academizer", version=SERVICE_VERSION)

_cors_raw = os.environ.get("PPT_ACADEMIZER_CORS_ORIGINS", "").strip()
if _cors_raw:
    _origins = [o.strip() for o in _cors_raw.split(",") if o.strip()]
    app.add_middleware(
        CORSMiddleware,
        allow_origins=_origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
        expose_headers=[
            "X-Academize-Warnings",
            "X-Academize-Slide-Count",
            "X-Academize-Profile",
            "X-Academize-Pipeline",
            "X-Academize-Source-Format",
            "Content-Disposition",
        ],
    )

if WEB.is_dir():
    app.mount("/static", StaticFiles(directory=WEB), name="static")

T = TypeVar("T")


async def _run_blocking(fn: Callable[..., T], /, *args, **kwargs) -> T:
    return await asyncio.to_thread(fn, *args, **kwargs)


def _schedule_tmp_cleanup(background_tasks: BackgroundTasks, tmp: Path) -> None:
    background_tasks.add_task(remove_upload_tmp, tmp)


def _academize_file_response(
    out_path: Path,
    *,
    warnings: list,
    slide_count: int,
    meta: dict,
    background_tasks: BackgroundTasks,
    work_tmp: Path,
) -> FileResponse:
    warn_payload = {
        "slide_count": slide_count,
        "warnings": warnings,
        "profile": meta.get("route_profile"),
        "pipeline": meta.get("pipeline"),
        "deck_kind": meta.get("deck_kind"),
        "service_version": meta.get("service_version"),
        "source_format": meta.get("source_format", "pptx"),
    }
    warn_header = base64.b64encode(
        json.dumps(warn_payload, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    _schedule_tmp_cleanup(background_tasks, work_tmp)
    return FileResponse(
        path=out_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=out_path.name,
        headers={
            "X-Academize-Warnings": warn_header,
            "X-Academize-Slide-Count": str(slide_count),
            "X-Academize-Profile": str(meta.get("route_profile", "")),
            "X-Academize-Pipeline": str(meta.get("pipeline", "")),
            "X-Academize-Source-Format": "pptx",
        },
    )


@app.get("/health")
def health():
    from scripts.academy_template import resolve_academy_template_path
    from scripts.migrate_version import MIGRATE_ENGINE_VERSION

    try:
        resolve_academy_template_path()
        tpl_ok = True
    except Exception:
        tpl_ok = False
    return {
        "ok": True,
        "service": "ppt-academizer",
        "service_version": SERVICE_VERSION,
        "migrate_engine_version": MIGRATE_ENGINE_VERSION,
        "engine_root": str(ENGINE_ROOT),
        "latest_release_doc": LATEST_RELEASE_DOC,
        "template_configured": tpl_ok,
        "max_upload_mb": max_upload_mb(),
        "max_upload_bytes": max_upload_bytes(),
        "standard_max_slides": standard_max_slides(),
        "supported_upload": [".pptx"],
        "skip_pp_repair": __import__("os").environ.get("PPT_ACADEMIZER_SKIP_PP_REPAIR", ""),
    }


@app.get("/config.js")
def config_js():
    path = WEB / "config.js"
    if not path.is_file():
        raise HTTPException(status_code=404, detail="web/config.js missing")
    return FileResponse(path, media_type="application/javascript")


@app.get("/", response_class=HTMLResponse)
def index():
    index_path = WEB / "index.html"
    if not index_path.is_file():
        raise HTTPException(status_code=404, detail="web/index.html missing")
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


@app.post("/wizard/preview")
async def wizard_preview(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    deck_title: str = Form(""),
    deck_subtitle: str = Form("PPT 아카데미화"),
):
    """Step 1: deck profile cards + pipeline-specific slide previews."""
    tmp: Path | None = None
    scheduled_cleanup = False
    try:
        content, filename = await read_upload_bounded(file)
        tmp, src = write_upload_tmp(content, filename)
        from core.wizard_preview import build_wizard_preview

        title = deck_title.strip() or None
        subtitle = deck_subtitle.strip() or "PPT 아카데미화"
        payload = await _run_blocking(
            build_wizard_preview, src, deck_title=title, deck_subtitle=subtitle
        )
        _schedule_tmp_cleanup(background_tasks, tmp)
        scheduled_cleanup = True
        return JSONResponse(payload)
    except FileNotFoundError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except HTTPException:
        raise
    except Exception as e:
        raise internal_server_error(e) from e
    finally:
        if tmp is not None and not scheduled_cleanup:
            remove_upload_tmp(tmp)


@app.post("/analyze")
async def analyze(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    profile: str = Form("auto"),
    deck_title: str = Form(""),
    deck_subtitle: str = Form("PPT 아카데미화"),
):
    tmp: Path | None = None
    scheduled_cleanup = False
    try:
        content, filename = await read_upload_bounded(file)
        tmp, src = write_upload_tmp(content, filename)
        from core.analyze import analyze_presentation

        title = deck_title.strip() or None
        subtitle = deck_subtitle.strip() or "PPT 아카데미화"
        payload = await _run_blocking(
            analyze_presentation,
            src,
            deck_title=title,
            deck_subtitle=subtitle,
            profile=profile,
        )
        _schedule_tmp_cleanup(background_tasks, tmp)
        scheduled_cleanup = True
        return JSONResponse(payload)
    except FileNotFoundError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except HTTPException:
        raise
    except Exception as e:
        raise internal_server_error(e) from e
    finally:
        if tmp is not None and not scheduled_cleanup:
            remove_upload_tmp(tmp)


@app.post("/academize")
async def academize(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    profile: str = Form("auto"),
    quality_mode: str = Form("standard"),
    estimated_output_slides: str = Form(""),
    deck_title: str = Form(""),
    deck_subtitle: str = Form("PPT 아카데미화"),
):
    tmp: Path | None = None
    scheduled_cleanup = False
    try:
        content, filename = await read_upload_bounded(file)
        tmp, src = write_upload_tmp(content, filename)
        from core.pipeline import academize_pptx

        title = deck_title.strip() or None
        subtitle = deck_subtitle.strip() or "PPT 아카데미화"
        est_out: int | None = None
        if estimated_output_slides.strip().isdigit():
            est_out = int(estimated_output_slides.strip())

        out_path, warnings, slide_count, meta = await _run_blocking(
            academize_pptx,
            src,
            work_dir=tmp,
            deck_title=title,
            deck_subtitle=subtitle,
            profile=profile,
            quality_mode=quality_mode,
            estimated_output_slides=est_out,
        )
        scheduled_cleanup = True
        return _academize_file_response(
            out_path,
            warnings=warnings,
            slide_count=slide_count,
            meta=meta,
            background_tasks=background_tasks,
            work_tmp=tmp,
        )
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=503,
            detail=f"아카데미 템플릿을 찾을 수 없습니다. TEMPLATE_PPTX를 설정하세요. ({e})",
        ) from e
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except HTTPException:
        raise
    except Exception as e:
        raise internal_server_error(e) from e
    finally:
        if tmp is not None and not scheduled_cleanup:
            remove_upload_tmp(tmp)
