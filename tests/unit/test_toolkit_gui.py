from pathlib import Path
from types import SimpleNamespace
import logging
import tkinter as tk

import pytest

from excel_splitter.controller import AppController


@pytest.fixture
def toolkit(monkeypatch, tk_root):
    from excel_splitter.toolkit_gui import ExcelFileToolkitGui

    monkeypatch.setattr("excel_splitter.gui.configure_logging", lambda: logging.getLogger("toolkit-test"))
    root = tk_root
    calls = []

    def preview(sources, target):
        calls.append((sources, target))
        return SimpleNamespace(
            inputs=tuple(SimpleNamespace(source=p, sheet_name="Data", row_count=2) for p in sources),
            target=target, prior_signature=None, row_count=2 * len(sources),
        )

    def execute(plan, overwrite, progress):
        calls.append((plan, overwrite))
        progress(2, 2, "done")
        return plan.target

    gui = ExcelFileToolkitGui(
        root, AppController(SimpleNamespace(shutdown=lambda: None)),
        SimpleNamespace(preview=preview, execute=execute),
    )
    gui.notebook.select(1)
    yield gui, calls
    for pending in root.tk.call("after", "info"):
        root.after_cancel(pending)
    for child in root.winfo_children():
        child.destroy()


@pytest.fixture(scope="module")
def tk_root():
    root = tk.Tk()
    root.withdraw()
    yield root
    root.destroy()


def complete_worker(gui):
    event = gui.events.get(timeout=3)
    while event[0] == "progress":
        gui._show_progress(*event[1:])
        event = gui.events.get(timeout=3)
    assert event[0] == "ok", event
    gui._handle_ok(event[1])


def add_files(gui, monkeypatch):
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilenames", lambda **_: ("b.xlsx", "a.xlsx", "b.xlsx"))
    gui._add_merge_files()
    gui.merge_output_var.set(str(Path("merged.xlsx").resolve()))


def test_file_order_preview_and_changes_invalidate_merge(toolkit, monkeypatch):
    gui, calls = toolkit
    add_files(gui, monkeypatch)
    assert len(gui.merge_sources) == 2
    gui.merge_tree.selection_set("1")
    gui._move_merge_file(-1)
    assert [p.name for p in gui.merge_sources] == ["a.xlsx", "b.xlsx"]
    gui._preview_merge()
    complete_worker(gui)
    assert calls[0][0] == tuple(gui.merge_sources)
    assert gui.merge_preview.row_count == 4
    assert str(gui.merge_button["state"]) == "normal"
    gui.merge_tree.selection_set("0")
    gui._remove_merge_file()
    assert gui.merge_preview is None
    assert str(gui.merge_button["state"]) == "disabled"


def test_running_merge_blocks_split_and_close_then_restores(toolkit, monkeypatch):
    gui, _ = toolkit
    add_files(gui, monkeypatch)
    gui._preview_merge()
    assert gui.notebook.tab(0, "state") == "disabled"
    assert str(gui.merge_preview_button["state"]) == "disabled"
    gui._on_close()
    assert gui.root.winfo_exists()
    complete_worker(gui)
    assert gui.notebook.tab(0, "state") == "normal"
    assert str(gui.merge_preview_button["state"]) == "normal"
    gui.notebook.select(0)
    gui._set_busy(True)
    assert gui.notebook.tab(1, "state") == "disabled"
    gui._set_busy(False)
    assert gui.notebook.tab(1, "state") == "normal"


def test_merge_requires_confirmation_and_consumes_preview(toolkit, monkeypatch):
    gui, calls = toolkit
    add_files(gui, monkeypatch)
    gui._preview_merge()
    complete_worker(gui)
    plan = gui.merge_preview
    plan.prior_signature = object()
    prompts = []
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.askyesno", lambda _, text, **__: prompts.append(text) or False)
    gui._merge()
    assert len(calls) == 1
    assert "덮어" in prompts[0]
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.askyesno", lambda *_, **__: True)
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda *_, **__: None)
    gui._merge()
    complete_worker(gui)
    assert calls[-1] == (plan, True)
    assert gui.merge_preview is None
    assert str(gui.merge_button["state"]) == "disabled"


def test_merge_error_clears_stale_preview_and_restores_controls(toolkit, monkeypatch):
    from excel_splitter.errors import WorkbookValidationError

    gui, _ = toolkit
    add_files(gui, monkeypatch)
    gui._preview_merge()
    complete_worker(gui)
    gui._set_busy(True)
    monkeypatch.setattr("excel_splitter.gui.messagebox.showerror", lambda *_, **__: None)
    gui._handle_error(WorkbookValidationError("changed input"))
    assert gui.merge_preview is None
    assert gui.notebook.tab(0, "state") == "normal"
    assert str(gui.merge_preview_button["state"]) == "normal"
