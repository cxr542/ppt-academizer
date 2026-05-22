#!/usr/bin/env python3
"""Pilot: compare spec vs migrate_cmp outputs for a source deck (plan pilot-cmp-compare).

Usage:
  cd apps/ppt-academizer
  export TEMPLATE_PPTX="…"
  ../ppt-test/.venv/bin/python scripts/compare_deck_paths.py --source /path/to/deck.pptx
  ../ppt-test/.venv/bin/python scripts/compare_deck_paths.py --source tests/fixtures/cmp-like-partner.pptx
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.ppt_test_path import ensure_engine_on_path  # noqa: E402

ensure_engine_on_path()
sys.path.insert(0, str(ROOT))

from pptx import Presentation  # noqa: E402

from core.analyze import analyze_presentation  # noqa: E402
from core.pipeline import academize_pptx  # noqa: E402
from scripts.deck_profile import detect_deck_profile  # noqa: E402


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def compare(source: Path, out_dir: Path) -> dict:
    source = source.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    report: dict = {"source": str(source), "paths": {}}

    prs = Presentation(str(source))
    profile, meta = detect_deck_profile(prs, filename_hint=source.name)
    report["detected_profile"] = profile
    report["profile_meta"] = meta
    report["source_slides"] = len(prs.slides)

    analysis = analyze_presentation(source)
    report["spec_preview"] = {
        "route": analysis["route_profile"],
        "spec_slides": analysis["spec_slide_count"],
        "validation_errors": analysis["validation_errors"],
    }
    (out_dir / f"analyze-{_stamp()}.json").write_text(
        json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    for route in ("spec", "migrate_cmp"):
        try:
            path, warnings, count, meta = academize_pptx(
                source, work_dir=out_dir / f"work-{route}", profile=route
            )
            report["paths"][route] = {
                "output": str(path),
                "slides": count,
                "pipeline": meta.get("pipeline"),
                "warnings": len(warnings),
            }
        except Exception as exc:
            report["paths"][route] = {"error": repr(exc)}

    report_path = out_dir / f"compare-report-{_stamp()}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    report["report_file"] = str(report_path)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare academize pipelines")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, default=ROOT / "output" / "compare")
    args = parser.parse_args()

    if not args.source.is_file():
        print(f"Not found: {args.source}", file=sys.stderr)
        return 2

    report = compare(args.source, args.out_dir)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
