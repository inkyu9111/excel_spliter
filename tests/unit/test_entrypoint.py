import builtins
import sys

import pytest

from excel_splitter import __main__ as entrypoint


def test_pywin32_self_test_imports_without_starting_gui(monkeypatch) -> None:
    imported: list[str] = []
    original_import = builtins.__import__

    def track_import(name, *args, **kwargs):
        imported.append(name)
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(sys, "argv", ["ExcelSplitter.exe", "--self-test-pywin32"])
    monkeypatch.setattr(builtins, "__import__", track_import)
    monkeypatch.setattr(
        entrypoint,
        "main",
        lambda: (_ for _ in ()).throw(AssertionError("GUI must not start")),
    )

    assert entrypoint.run() == 0
    assert "pythoncom" in imported
    assert "win32com.client" in imported


def test_pywin32_self_test_propagates_import_failure(monkeypatch) -> None:
    original_import = builtins.__import__

    def fail_pythoncom(name, *args, **kwargs):
        if name == "pythoncom":
            raise ImportError("missing pythoncom")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(sys, "argv", ["ExcelSplitter.exe", "--self-test-pywin32"])
    monkeypatch.setattr(builtins, "__import__", fail_pythoncom)
    monkeypatch.setattr(
        entrypoint,
        "main",
        lambda: (_ for _ in ()).throw(AssertionError("GUI must not start")),
    )

    with pytest.raises(ImportError, match="missing pythoncom"):
        entrypoint.run()


def test_normal_entrypoint_starts_gui(monkeypatch) -> None:
    called: list[bool] = []
    monkeypatch.setattr(sys, "argv", ["ExcelSplitter.exe"])
    monkeypatch.setattr(entrypoint, "main", lambda: called.append(True))

    assert entrypoint.run() == 0
    assert called == [True]
