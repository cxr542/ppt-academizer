"""Run each conversion in its own short-lived process.

Why: a deck conversion allocates hundreds of MB of native memory (lxml trees, image buffers) that
the C allocator does not hand back to the OS after the request. In one long-lived server process
that memory piles up request after request until the 512MB instance is killed (OOM). A worker
process that exits after one job returns everything, and one job at a time bounds the peak.

``PPT_ACADEMIZER_ISOLATE=0`` runs jobs in a thread of the server process instead (tests/debug).
``PPT_ACADEMIZER_MAX_CONCURRENT`` (default 1) caps simultaneous conversions.
"""

from __future__ import annotations

import asyncio
import functools
import multiprocessing
import os
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from typing import Callable, TypeVar

T = TypeVar("T")

_semaphore: asyncio.Semaphore | None = None
_semaphore_loop: asyncio.AbstractEventLoop | None = None


class ConversionWorkerDied(RuntimeError):
    """The worker process exited abnormally (typically killed for exceeding memory)."""


def isolation_enabled() -> bool:
    return os.environ.get("PPT_ACADEMIZER_ISOLATE", "1").strip().lower() not in ("0", "false", "no", "off")


def max_concurrent() -> int:
    raw = os.environ.get("PPT_ACADEMIZER_MAX_CONCURRENT", "").strip()
    return max(1, int(raw)) if raw.isdigit() else 1


def _slots() -> asyncio.Semaphore:
    global _semaphore, _semaphore_loop
    loop = asyncio.get_running_loop()
    if _semaphore is None or _semaphore_loop is not loop:
        _semaphore, _semaphore_loop = asyncio.Semaphore(max_concurrent()), loop
    return _semaphore


async def run_isolated(fn: Callable[..., T], /, *args, **kwargs) -> T:
    if not isolation_enabled():
        return await asyncio.to_thread(fn, *args, **kwargs)
    async with _slots():
        loop = asyncio.get_running_loop()
        # "spawn": a fresh interpreter, not a fork of the (already large) server process.
        executor = ProcessPoolExecutor(max_workers=1, mp_context=multiprocessing.get_context("spawn"))
        try:
            return await loop.run_in_executor(executor, functools.partial(fn, *args, **kwargs))
        except BrokenProcessPool as exc:
            raise ConversionWorkerDied(str(exc)) from exc
        finally:
            executor.shutdown(wait=False, cancel_futures=True)
