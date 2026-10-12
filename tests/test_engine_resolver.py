"""The engine resolver must never fall back to the retired ppt-test copy."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.ppt_test_path import resolve_ppt_test_root


@pytest.fixture(autouse=True)
def _no_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("PPT_ENGINE_ROOT", raising=False)
    monkeypatch.delenv("PPT_TEST_ROOT", raising=False)


def test_bundled_engine_is_used(tmp_path: Path) -> None:
    (tmp_path / "engine" / "scripts").mkdir(parents=True)
    (tmp_path / "engine" / "scripts" / "academy_deck_build_lib.py").write_text("", encoding="utf-8")
    assert resolve_ppt_test_root(tmp_path) == (tmp_path / "engine").resolve()


def test_missing_engine_raises_instead_of_falling_back(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError, match="git checkout -- engine"):
        resolve_ppt_test_root(tmp_path)


def test_explicit_override_wins(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    other = tmp_path / "custom-engine"
    other.mkdir()
    monkeypatch.setenv("PPT_ENGINE_ROOT", str(other))
    assert resolve_ppt_test_root(tmp_path) == other.resolve()


def test_real_repo_resolves_to_its_own_engine() -> None:
    root = resolve_ppt_test_root()
    assert root.name == "engine" and (root / "scripts" / "academy_deck_build_lib.py").is_file()
