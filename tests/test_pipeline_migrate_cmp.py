from __future__ import annotations

from pathlib import Path

from core import pipeline


def test_migrate_cmp_retries_without_ooxml_repair_on_import_error(
    monkeypatch,
    tmp_path: Path,
) -> None:
    import scripts.academy_deck_build_lib as deck_lib
    import scripts.build_cmp_academy as cmp_builder

    calls: list[bool] = []
    source = tmp_path / "source.pptx"
    source.write_bytes(b"source")

    def fake_migrate_cmp_deck(
        source_path: Path,
        output_path: Path,
        *,
        skip_ooxml_repair: bool,
    ) -> tuple[Path, list[dict], int]:
        calls.append(skip_ooxml_repair)
        if not skip_ooxml_repair:
            raise ImportError("dlopen(pyexpat): Symbol not found")
        output_path.write_bytes(source_path.read_bytes())
        return output_path, [{"code": "MIGRATE_META", "deck_kind": "cmp"}], 9

    monkeypatch.setattr(cmp_builder, "migrate_cmp_deck", fake_migrate_cmp_deck)
    monkeypatch.setattr(deck_lib, "finalize_academy_package", lambda _output: 0)
    monkeypatch.setattr(pipeline, "maybe_mac_powerpoint_repair", lambda _output: [])

    output, warnings, slide_count, meta = pipeline._academize_migrate_cmp(
        source,
        tmp=tmp_path,
    )

    assert calls == [False, True]
    assert output.read_bytes() == b"source"
    assert slide_count == 9
    assert meta["deck_kind"] == "cmp"
    assert any(w["code"] == "OOXML_REPAIR_FAILED" for w in warnings)


def test_validator_subprocess_gets_the_engine_on_pythonpath(monkeypatch) -> None:
    """The validator imports scripts.safe_xml; without PYTHONPATH it failed on every local run."""
    import os

    from scripts.build_cmp_academy import PPT_TEST, validator_subprocess_env

    monkeypatch.delenv("PYTHONPATH", raising=False)
    assert validator_subprocess_env()["PYTHONPATH"].split(os.pathsep)[0] == str(PPT_TEST)

    monkeypatch.setenv("PYTHONPATH", "/somewhere/else")
    parts = validator_subprocess_env()["PYTHONPATH"].split(os.pathsep)
    assert parts[0] == str(PPT_TEST) and "/somewhere/else" in parts


def test_validator_runs_when_given_that_environment(tmp_path: Path) -> None:
    """End to end: the script that used to die with ModuleNotFoundError now validates a deck."""
    import subprocess
    import sys

    from pptx import Presentation

    from scripts.build_cmp_academy import PPT_TEST, validator_subprocess_env

    deck = tmp_path / "ok.pptx"
    Presentation().save(deck)
    proc = subprocess.run(
        [sys.executable, str(PPT_TEST / "scripts" / "validate_pptx.py"), str(deck)],
        capture_output=True,
        text=True,
        env=validator_subprocess_env(),
    )
    assert "ModuleNotFoundError" not in (proc.stderr + proc.stdout), proc.stderr[-400:]
