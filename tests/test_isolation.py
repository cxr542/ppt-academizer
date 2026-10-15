"""Conversions run in a throw-away worker process so memory is returned after each job."""

from __future__ import annotations

import asyncio
import os

import pytest

from api.isolation import ConversionWorkerDied, max_concurrent, run_isolated


def test_job_runs_in_a_different_process(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_ISOLATE", raising=False)
    pid = asyncio.run(run_isolated(os.getpid))
    assert pid != os.getpid()


def test_each_job_gets_a_fresh_process(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_ISOLATE", raising=False)

    async def two():
        return await run_isolated(os.getpid), await run_isolated(os.getpid)

    first, second = asyncio.run(two())
    assert first != second


def test_exceptions_keep_their_type(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_ISOLATE", raising=False)
    with pytest.raises(ValueError):
        asyncio.run(run_isolated(int, "not a number"))


def test_a_killed_worker_becomes_conversion_worker_died(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_ISOLATE", raising=False)
    with pytest.raises(ConversionWorkerDied):
        asyncio.run(run_isolated(os._exit, 1))
    # the next job is unaffected (a new executor is used per job)
    assert asyncio.run(run_isolated(int, "7")) == 7


def test_isolation_can_be_switched_off(monkeypatch) -> None:
    monkeypatch.setenv("PPT_ACADEMIZER_ISOLATE", "0")
    assert asyncio.run(run_isolated(os.getpid)) == os.getpid()


def test_concurrency_defaults_to_one_and_is_configurable(monkeypatch) -> None:
    monkeypatch.delenv("PPT_ACADEMIZER_MAX_CONCURRENT", raising=False)
    assert max_concurrent() == 1
    monkeypatch.setenv("PPT_ACADEMIZER_MAX_CONCURRENT", "3")
    assert max_concurrent() == 3
    monkeypatch.setenv("PPT_ACADEMIZER_MAX_CONCURRENT", "0")
    assert max_concurrent() == 1
