"""sync_engine_from_ppt_test.py must never silently replace a maintained engine/."""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "sync_engine_from_ppt_test.py"


@pytest.fixture()
def sync(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    spec = importlib.util.spec_from_file_location("sync_engine_under_test", SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)

    root = tmp_path / "app"
    (root / "engine" / "scripts").mkdir(parents=True)
    (root / "engine" / "scripts" / "mine.py").write_text("FIXED = True\n", encoding="utf-8")
    monkeypatch.setattr(mod, "ROOT", root)
    monkeypatch.setattr(mod, "ENGINE", root / "engine")

    ppt_test = tmp_path / "ppt-test"
    (ppt_test / "scripts").mkdir(parents=True)
    (ppt_test / "scripts" / "mine.py").write_text("FIXED = False  # old copy\n", encoding="utf-8")
    monkeypatch.setenv("PPT_TEST_ROOT", str(ppt_test))
    return mod, root


def test_refuses_without_force_and_leaves_engine_untouched(sync, capsys) -> None:
    mod, root = sync
    assert mod.main([]) == 3
    assert (root / "engine" / "scripts" / "mine.py").read_text(encoding="utf-8") == "FIXED = True\n"
    assert not list(root.glob("engine.bak-*"))
    assert "--force" in capsys.readouterr().err


def test_force_moves_engine_to_backup_instead_of_deleting(sync) -> None:
    mod, root = sync
    assert mod.main(["--force"]) == 0
    backups = list(root.glob("engine.bak-*"))
    assert len(backups) == 1
    assert (backups[0] / "scripts" / "mine.py").read_text(encoding="utf-8") == "FIXED = True\n"
    assert "old copy" in (root / "engine" / "scripts" / "mine.py").read_text(encoding="utf-8")


def test_missing_engine_syncs_without_force(sync, tmp_path: Path) -> None:
    mod, root = sync
    import shutil

    shutil.rmtree(root / "engine")
    assert mod.main([]) == 0
    assert (root / "engine" / "scripts" / "mine.py").is_file()
    assert not list(root.glob("engine.bak-*"))


def test_missing_source_is_an_error_not_a_wipe(sync, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    mod, root = sync
    monkeypatch.setenv("PPT_TEST_ROOT", str(tmp_path / "does-not-exist"))
    assert mod.main(["--force"]) == 2
    assert (root / "engine" / "scripts" / "mine.py").is_file()
