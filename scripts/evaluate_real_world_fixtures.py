#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = [
#   "python-pptx>=1.0.0",
#   "Pillow>=10.0.0",
# ]
# ///
"""Evaluate real-world PPT fixtures for basic academizer smoke signals.

# ─── How to run ───
# cd ppt-academizer
# .venv/bin/python scripts/evaluate_real_world_fixtures.py
# .venv/bin/python scripts/evaluate_real_world_fixtures.py --metadata-only
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from zipfile import BadZipFile

ROOT = Path(__file__).resolve().parent.parent
os.environ.setdefault("PPT_ACADEMIZER_SKIP_PP_REPAIR", "1")
sys.path.insert(0, str(ROOT))

from core.ppt_test_path import ensure_engine_on_path  # noqa: E402

ensure_engine_on_path()

from core.pipeline import academize_pptx  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.exc import PackageNotFoundError  # noqa: E402

FIXTURE_DIR = ROOT / "tests" / "fixtures" / "real_world"
OUTPUT_DIR = ROOT / "outputs" / "evaluation"


@dataclass(frozen=True)
class Fixture:
    priority: int
    filename: str
    original_name: str
    purpose: str

    @property
    def path(self) -> Path:
        return FIXTURE_DIR / self.filename


@dataclass(frozen=True)
class FixtureResult:
    priority: int
    fixture: str
    original_name: str
    purpose: str
    source_exists: bool
    source_size_bytes: int | None
    source_opens: bool
    source_slide_count: int | None
    conversion_attempted: bool
    conversion_ok: bool
    output_path: str | None
    output_size_bytes: int | None
    output_opens: bool
    output_slide_count: int | None
    warning_count: int | None
    warnings: list[str]
    error: str | None


FIXTURES: tuple[Fixture, ...] = (
    Fixture(
        1,
        "01_k8s_dashboard_lab_lecture.pptx",
        "k8s_dashboard_lab_lecture.pptx",
        "실습 강의안형: 학습 목표, 실습 흐름, YAML, 체크리스트, 강사용 멘트 확인",
    ),
    Fixture(
        2,
        "02_cmp_core_technology.pptx",
        "클라우드 구현기술(CMP)_v1.0_수정요청.pptx",
        "기술 개념 설명형: 개념 정의, 비유 설명, 기대효과 표 구조 확인",
    ),
    Fixture(
        3,
        "03_vmware_winback_strategy_report.pptx",
        "오케스트로 VMware 윈백 시장 주도 전략 보고.pptx",
        "전략 보고서형: 큰 제목, 숫자 지표, 로드맵, 전략 메시지 확인",
    ),
    Fixture(
        4,
        "04_academy_registration_page_plan.pptx",
        "[기획안] 오케스트로 아카데미 교육신청 페이지_260128 v1.pptx",
        "기획/요구사항형: 긴 텍스트, 표, CTA, 관리자 요구사항, 폼 문구 확인",
    ),
    Fixture(
        5,
        "05_contrabass_base_technology.pptx",
        "CONTRABASS 기반기술@260504_수정 요청.pptx",
        "복합 기술 발표형: 다중 기술 섹션, 긴 자료, 표, 이미지 출처 확인",
    ),
)


def _slide_count(path: Path) -> int:
    return len(Presentation(str(path)).slides)


def _warning_messages(warnings: list[dict]) -> list[str]:
    messages: list[str] = []
    for warning in warnings:
        message = warning.get("message")
        code = warning.get("code")
        if isinstance(message, str):
            messages.append(message)
        elif isinstance(code, str):
            messages.append(code)
        else:
            messages.append(str(warning))
    return messages


def _result_to_dict(result: FixtureResult) -> dict[str, str | int | bool | None | list[str]]:
    return {
        "priority": result.priority,
        "fixture": result.fixture,
        "original_name": result.original_name,
        "purpose": result.purpose,
        "source_exists": result.source_exists,
        "source_size_bytes": result.source_size_bytes,
        "source_opens": result.source_opens,
        "source_slide_count": result.source_slide_count,
        "conversion_attempted": result.conversion_attempted,
        "conversion_ok": result.conversion_ok,
        "output_path": result.output_path,
        "output_size_bytes": result.output_size_bytes,
        "output_opens": result.output_opens,
        "output_slide_count": result.output_slide_count,
        "warning_count": result.warning_count,
        "warnings": result.warnings,
        "error": result.error,
    }


def _evaluate_fixture(fixture: Fixture, *, convert: bool) -> FixtureResult:
    path = fixture.path
    if not path.is_file():
        return FixtureResult(
            fixture.priority,
            fixture.filename,
            fixture.original_name,
            fixture.purpose,
            False,
            None,
            False,
            None,
            convert,
            False,
            None,
            None,
            False,
            None,
            None,
            [],
            f"missing fixture: {path}",
        )

    try:
        source_slide_count = _slide_count(path)
    except (BadZipFile, KeyError, PackageNotFoundError, OSError) as err:
        return FixtureResult(
            fixture.priority,
            fixture.filename,
            fixture.original_name,
            fixture.purpose,
            True,
            path.stat().st_size,
            False,
            None,
            convert,
            False,
            None,
            None,
            False,
            None,
            None,
            [],
            f"source open failed: {err}",
        )

    if not convert:
        return FixtureResult(
            fixture.priority,
            fixture.filename,
            fixture.original_name,
            fixture.purpose,
            True,
            path.stat().st_size,
            True,
            source_slide_count,
            False,
            False,
            None,
            None,
            False,
            None,
            None,
            [],
            None,
        )

    work_dir = OUTPUT_DIR / "artifacts" / path.stem
    try:
        output, warnings, _, _ = academize_pptx(
            path,
            deck_title=path.stem,
            deck_subtitle="Real-world fixture evaluation",
            work_dir=work_dir,
            quality_mode="unlimited",
        )
        output_slide_count = _slide_count(output)
        return FixtureResult(
            fixture.priority,
            fixture.filename,
            fixture.original_name,
            fixture.purpose,
            True,
            path.stat().st_size,
            True,
            source_slide_count,
            True,
            True,
            str(output.relative_to(ROOT)),
            output.stat().st_size,
            True,
            output_slide_count,
            len(warnings),
            _warning_messages(warnings),
            None,
        )
    except (
        BadZipFile,
        FileNotFoundError,
        KeyError,
        OSError,
        PackageNotFoundError,
        RuntimeError,
        ValueError,
    ) as err:
        return FixtureResult(
            fixture.priority,
            fixture.filename,
            fixture.original_name,
            fixture.purpose,
            True,
            path.stat().st_size,
            True,
            source_slide_count,
            True,
            False,
            None,
            None,
            False,
            None,
            None,
            [],
            f"conversion failed: {err}",
        )


def _write_json(results: list[FixtureResult]) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / "real_world_results.json"
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "fixture_dir": str(FIXTURE_DIR.relative_to(ROOT)),
        "results": [_result_to_dict(result) for result in results],
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def _write_markdown(results: list[FixtureResult]) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / "real_world_results.md"
    lines = [
        "# Real-world fixture smoke results",
        "",
        f"Generated: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "| Priority | Fixture | Source slides | Conversion | Output slides | Warnings | Error |",
        "|---:|---|---:|---|---:|---:|---|",
    ]
    for result in results:
        conversion = "ok" if result.conversion_ok else "not run"
        if result.conversion_attempted and not result.conversion_ok:
            conversion = "failed"
        lines.append(
            "| "
            f"{result.priority} | `{result.fixture}` | "
            f"{result.source_slide_count if result.source_slide_count is not None else ''} | "
            f"{conversion} | "
            f"{result.output_slide_count if result.output_slide_count is not None else ''} | "
            f"{result.warning_count if result.warning_count is not None else ''} | "
            f"{result.error or ''} |"
        )
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Evaluate real-world PPT fixtures")
    parser.add_argument(
        "--metadata-only",
        action="store_true",
        help="Only verify source PPTX files without running academize conversion.",
    )
    args = parser.parse_args()

    results = [_evaluate_fixture(fixture, convert=not args.metadata_only) for fixture in FIXTURES]
    json_path = _write_json(results)
    md_path = _write_markdown(results)

    print(f"Wrote {json_path.relative_to(ROOT)}")
    print(f"Wrote {md_path.relative_to(ROOT)}")
    failed = [result for result in results if result.error]
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
