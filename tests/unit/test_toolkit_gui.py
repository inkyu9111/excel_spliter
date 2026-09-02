from pathlib import Path
from types import SimpleNamespace
from dataclasses import replace
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


def test_compare_selection_creates_unused_output_and_saves_result(toolkit, monkeypatch, tmp_path):
    gui, _ = toolkit
    gui.notebook.select(2)
    reference, comparison = tmp_path / "reference.xlsx", tmp_path / "comparison.xlsx"
    (tmp_path / "comparison_비교결과.xlsx").touch()
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilename", lambda **_: str(reference))
    gui._browse_compare_input("reference")
    assert str(gui.compare_button["state"]) == "disabled"
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilename", lambda **_: str(comparison))
    gui._browse_compare_input("comparison")
    output = tmp_path / "comparison_비교결과 (2).xlsx"
    assert Path(gui.compare_output_var.get()) == output
    assert str(gui.compare_button["state"]) == "normal"

    def execute(actual_reference, actual_comparison, target, *, progress):
        assert (actual_reference, actual_comparison, target) == (reference, comparison, output)
        progress(1, 1, "Data")
        return SimpleNamespace(target=target, changed_cells=3, missing_sheets=("누락 시트",),
                               missing_rows=0, missing_columns=())

    gui.compare_service = SimpleNamespace(execute=execute)
    messages = []
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda _, text, **__: messages.append(text))
    gui._compare()
    assert gui.notebook.tab(0, "state") == gui.notebook.tab(1, "state") == "disabled"
    assert str(gui.compare_button["state"]) == "disabled"
    complete_worker(gui)
    assert str(output) in messages[0] and "3" in messages[0] and "누락 시트" in messages[0]
    assert gui.notebook.tab(0, "state") == gui.notebook.tab(1, "state") == "normal"
    assert str(gui.compare_button["state"]) == "normal"


def test_compare_output_browse_and_error_restore_controls(toolkit, monkeypatch, tmp_path):
    from excel_splitter.errors import WorkbookValidationError

    gui, _ = toolkit
    gui.notebook.select(2)
    gui.compare_reference_var.set(str(tmp_path / "reference.xlsx"))
    gui.compare_comparison_var.set(str(tmp_path / "comparison.xlsx"))
    output = tmp_path / "custom.xlsx"
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.asksaveasfilename", lambda **_: str(output))
    gui._browse_compare_output()
    assert Path(gui.compare_output_var.get()) == output
    gui._set_busy(True)
    monkeypatch.setattr("excel_splitter.gui.messagebox.showerror", lambda *_, **__: None)
    gui._handle_error(WorkbookValidationError("existing output"))
    assert "오류" in gui.compare_status_var.get()
    assert not gui._busy and str(gui.compare_button["state"]) == "normal"


def test_key_compare_loads_tables_selects_multiple_columns_and_sends_options(toolkit, monkeypatch, tmp_path):
    gui, _ = toolkit
    gui.notebook.select(2)
    reference, comparison, target = (tmp_path / name for name in ("base.xlsx", "other.xlsx", "result.xlsx"))
    gui.compare_reference_var.set(str(reference))
    gui.compare_comparison_var.set(str(comparison))
    gui.compare_output_var.set(str(target))
    gui.compare_by_key_var.set(True)
    gui._compare_mode_changed()
    assert str(gui.compare_button["state"]) == "disabled"
    tables = (
        (SimpleNamespace(sheet_name="Base", table_name="BaseTable", columns=("b_col", "c_col", "amount")),),
        (SimpleNamespace(sheet_name="Other", table_name="OtherTable", columns=("amount", "c_col", "b_col")),),
    )

    def inspect_tables(actual_reference, actual_comparison):
        assert (actual_reference, actual_comparison) == (reference, comparison)
        return tables

    def execute(actual_reference, actual_comparison, actual_target, *, progress, **options):
        assert (actual_reference, actual_comparison, actual_target) == (reference, comparison, target)
        assert options == dict(key_columns=("b_col", "c_col"),
                               reference_table=("Base", "BaseTable"), comparison_table=("Other", "OtherTable"))
        return SimpleNamespace(target=target, changed_cells=2, missing_sheets=(), missing_rows=3,
                               missing_columns=("old_amount",))

    gui.compare_service = SimpleNamespace(inspect_tables=inspect_tables, execute=execute)
    gui._load_compare_tables()
    assert str(gui.compare_key_list["state"]) == "disabled"
    complete_worker(gui)
    assert gui.compare_key_list.get(0, "end") == ("b_col", "c_col", "amount")
    assert str(gui.compare_button["state"]) == "disabled"
    gui.compare_key_list.selection_set(0, 1)
    gui._render_compare_state()
    assert str(gui.compare_button["state"]) == "normal"
    messages = []
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda _, text, **__: messages.append(text))
    gui._compare()
    complete_worker(gui)
    assert "3" in messages[0] and "old_amount" in messages[0]


def test_key_compare_switching_table_or_file_clears_keys_and_position_mode_stays_available(toolkit, monkeypatch, tmp_path):
    gui, _ = toolkit
    gui.notebook.select(2)
    gui.compare_reference_var.set(str(tmp_path / "base.xlsx"))
    gui.compare_comparison_var.set(str(tmp_path / "other.xlsx"))
    gui.compare_output_var.set(str(tmp_path / "result.xlsx"))
    gui.compare_by_key_var.set(True)
    tables = (
        (SimpleNamespace(sheet_name="Base", table_name="Table1", columns=("id", "code")),),
        (SimpleNamespace(sheet_name="Other", table_name="Table2", columns=("id", "code")),
         SimpleNamespace(sheet_name="Other", table_name="Table3", columns=("code",))),
    )
    gui._handle_ok(("compare_tables", tables))
    gui.compare_key_list.selection_set(0)
    gui.compare_comparison_table_combo.current(1)
    gui._refresh_compare_keys()
    # Missing comparison columns remain selectable, so validation can name the missing key.
    assert gui.compare_key_list.get(0, "end") == ("id", "code")
    assert gui.compare_key_list.curselection() == ()
    assert str(gui.compare_button["state"]) == "disabled"
    gui.compare_key_list.selection_set(0)
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilename", lambda **_: str(tmp_path / "new.xlsx"))
    gui._browse_compare_input("reference")
    assert gui.compare_key_list.size() == 0
    assert str(gui.compare_button["state"]) == "disabled"
    gui.compare_by_key_var.set(False)
    gui._compare_mode_changed()
    assert str(gui.compare_button["state"]) == "normal"
    assert str(gui.compare_key_list["state"]) == "disabled"


def test_etc_loads_sheets_and_saves_selected_operations_to_new_file(toolkit, monkeypatch, tmp_path):
    gui, _ = toolkit
    gui.notebook.select(3)
    source = tmp_path / "source.xlsx"
    (tmp_path / "source_정리결과.xlsx").touch()
    output = tmp_path / "source_정리결과 (2).xlsx"
    calls = []

    def inspect_source(path):
        assert path == source
        return ("Keep", "Clean")

    def execute(path, sheet_name, target, *, remove_artifacts, reset_fill, progress):
        calls.append((path, sheet_name, target, remove_artifacts, reset_fill))
        progress(1, 1, sheet_name)
        return target

    gui.etc_service = SimpleNamespace(inspect_source=inspect_source, execute=execute)
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilename", lambda **_: str(source))
    gui._browse_etc_source()
    assert all(gui.notebook.tab(index, "state") == "disabled" for index in range(3))
    complete_worker(gui)
    assert tuple(gui.etc_sheet_combo["values"]) == ("Keep", "Clean")
    assert Path(gui.etc_output_var.get()) == output
    assert str(gui.etc_button["state"]) == "disabled"
    gui.etc_sheet_var.set("Clean")
    gui.etc_remove_artifacts_var.set(True)
    gui.etc_reset_fill_var.set(True)
    gui._render_etc_state()
    assert str(gui.etc_button["state"]) == "normal"
    messages = []
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda _, text, **__: messages.append(text))
    gui._run_etc()
    assert str(gui.etc_sheet_combo["state"]) == "disabled"
    assert str(gui.etc_button["state"]) == "disabled"
    complete_worker(gui)
    assert calls == [(source, "Clean", output, True, True)]
    assert str(output) in messages[0]
    assert all(gui.notebook.tab(index, "state") == "normal" for index in range(3))


def test_etc_options_work_independently_and_new_file_clears_stale_sheets(toolkit, monkeypatch, tmp_path):
    from excel_splitter.errors import WorkbookValidationError

    gui, _ = toolkit
    gui.notebook.select(3)
    gui.etc_source_var.set(str(tmp_path / "first.xlsx"))
    gui.etc_output_var.set(str(tmp_path / "result.xlsx"))
    gui._handle_ok(("etc_source", ("Old",)))
    calls = []

    def execute(source, sheet_name, target, **options):
        calls.append((options["remove_artifacts"], options["reset_fill"]))
        return target

    gui.etc_service = SimpleNamespace(execute=execute, inspect_source=lambda _: ("New",))
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda *_, **__: None)
    for remove, reset in ((True, False), (False, True)):
        gui.etc_remove_artifacts_var.set(remove)
        gui.etc_reset_fill_var.set(reset)
        gui._render_etc_state()
        assert str(gui.etc_button["state"]) == "normal"
        gui._run_etc()
        complete_worker(gui)
    assert calls == [(True, False), (False, True)]
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilename", lambda **_: str(tmp_path / "second.xlsx"))
    gui._browse_etc_source()
    assert gui.etc_sheet_var.get() == ""
    assert str(gui.etc_button["state"]) == "disabled"
    complete_worker(gui)
    assert tuple(gui.etc_sheet_combo["values"]) == ("New",)
    chosen = tmp_path / "custom.xlsx"
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.asksaveasfilename", lambda **_: str(chosen))
    gui._browse_etc_output()
    assert Path(gui.etc_output_var.get()) == chosen
    gui._set_busy(True)
    monkeypatch.setattr("excel_splitter.gui.messagebox.showerror", lambda *_, **__: None)
    gui._handle_error(WorkbookValidationError("protected sheet"))
    assert not gui._busy and "오류" in gui.etc_status_var.get()
    assert str(gui.etc_button["state"]) == "normal"


def test_output_entries_edit_state_without_key_events(toolkit, monkeypatch, tmp_path):
    gui, calls = toolkit
    split_output = tmp_path / "split-output"
    gui.controller.state = replace(
        gui.controller.state,
        source=tmp_path / "source.xlsx",
        sheet_name="Data",
        column_name="Team",
        output_dir=split_output,
        preview=object(),
    )
    gui._render_state(gui.controller.state)

    gui.output_entry.delete(0, "end")
    assert gui.output_var.get() == ""
    assert gui.controller.state.output_dir is None
    assert gui.controller.state.preview is None
    assert str(gui.preview_button["state"]) == "disabled"

    trailing_output = f"{split_output}\\"
    gui.output_entry.insert(0, trailing_output)
    assert gui.output_var.get() == trailing_output
    assert gui.controller.state.output_dir == split_output

    add_files(gui, monkeypatch)
    gui.merge_preview = SimpleNamespace()
    gui._render_merge_state()
    merge_output = tmp_path / "manual-merge.xlsx"
    gui.merge_output_entry.delete(0, "end")
    gui.merge_output_entry.insert(0, str(merge_output))
    assert gui.merge_preview is None
    gui._preview_merge()
    complete_worker(gui)
    assert calls[-1][1] == merge_output


def test_output_entries_update_readiness_and_busy_state(toolkit, tmp_path):
    gui, _ = toolkit
    assert str(gui.source_entry["state"]) == "readonly"
    assert all(str(entry["state"]) == "normal" for entry in (
        gui.output_entry, gui.merge_output_entry, gui.compare_output_entry, gui.etc_output_entry,
    ))

    gui.compare_reference_var.set(str(tmp_path / "reference.xlsx"))
    gui.compare_comparison_var.set(str(tmp_path / "comparison.xlsx"))
    gui.compare_output_entry.insert(0, str(tmp_path / "compare.xlsx"))
    assert str(gui.compare_button["state"]) == "normal"
    gui.compare_output_entry.delete(0, "end")
    assert str(gui.compare_button["state"]) == "disabled"

    gui.etc_source_var.set(str(tmp_path / "source.xlsx"))
    gui.etc_sheet_var.set("Data")
    gui.etc_remove_artifacts_var.set(True)
    gui.etc_output_entry.insert(0, str(tmp_path / "etc.xlsx"))
    assert str(gui.etc_button["state"]) == "normal"
    gui.etc_output_entry.delete(0, "end")
    assert str(gui.etc_button["state"]) == "disabled"

    gui._set_busy(True)
    assert all(str(entry["state"]) == "disabled" for entry in (
        gui.output_entry, gui.merge_output_entry, gui.compare_output_entry, gui.etc_output_entry,
    ))
    gui._set_busy(False)
    assert all(str(entry["state"]) == "normal" for entry in (
        gui.output_entry, gui.merge_output_entry, gui.compare_output_entry, gui.etc_output_entry,
    ))
    assert str(gui.source_entry["state"]) == "readonly"
