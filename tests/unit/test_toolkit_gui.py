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
    for bar in (gui.progress, gui.merge_progress, gui.compare_progress, gui.etc_progress):
        bar.stop()
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
    while gui._busy:
        event = gui.events.get(timeout=3)
        if event[0] == "progress":
            gui._show_progress(*event[1:])
        elif event[0] == "error":
            gui._handle_error(event[1])
        else:
            assert event[0] == "ok", event
            gui._handle_ok(event[1])


def add_files(gui, monkeypatch):
    monkeypatch.setattr("excel_splitter.toolkit_gui.filedialog.askopenfilenames", lambda **_: ("b.xlsx", "a.xlsx", "b.xlsx"))
    gui._add_merge_files()
    gui.merge_output_var.set(str(Path("merged.xlsx").resolve()))


def test_one_click_merge_preserves_order_without_manual_preview_or_confirmation(toolkit, monkeypatch):
    gui, calls = toolkit
    add_files(gui, monkeypatch)
    assert len(gui.merge_sources) == 2
    gui.merge_tree.selection_set("1")
    gui._move_merge_file(-1)
    assert [p.name for p in gui.merge_sources] == ["a.xlsx", "b.xlsx"]
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.askyesno", lambda *_, **__: pytest.fail("New output does not need confirmation"))
    gui.merge_button.invoke()
    complete_worker(gui)
    assert calls[0][0] == tuple(gui.merge_sources)
    assert calls[1][0].row_count == 4
    assert gui.result_paths["merge"] == (calls[0][1],)
    assert str(gui.merge_button["state"]) == "normal"
    gui.merge_tree.selection_set("0")
    gui._remove_merge_file()
    assert gui.merge_preview is None
    assert str(gui.merge_button["state"]) == "disabled"


def test_running_merge_blocks_split_and_close_then_restores(toolkit, monkeypatch):
    gui, _ = toolkit
    add_files(gui, monkeypatch)
    gui.merge_button.invoke()
    assert gui.notebook.tab(0, "state") == "disabled"
    assert str(gui.merge_button["state"]) == "disabled"
    gui._on_close()
    assert gui.root.winfo_exists()
    complete_worker(gui)
    assert gui.notebook.tab(0, "state") == "normal"
    assert str(gui.merge_button["state"]) == "normal"
    gui.notebook.select(0)
    gui._set_busy(True)
    assert gui.notebook.tab(1, "state") == "disabled"
    gui._set_busy(False)
    assert gui.notebook.tab(1, "state") == "normal"


def test_one_click_split_inspects_confirms_and_executes_current_settings(toolkit, monkeypatch, tmp_path):
    from excel_splitter.models import CanonicalKey, FileSignature, GroupSummary, OutputTarget, Preview, SplitResult, WorkbookSnapshot

    gui, _ = toolkit
    gui.notebook.select(0)
    key = CanonicalKey("text", "서울")
    source, target = tmp_path / "source.xlsx", tmp_path / "서울_분할.xlsx"
    snapshot = WorkbookSnapshot(source, FileSignature(1, 2, "sig"), "Data", "DataTable", "지역", 1,
                                (GroupSummary(key, "서울", 1, (1,)),))
    plan = Preview(snapshot, (OutputTarget(key, "서울", target, None),), ())
    stages = []

    def inspect(*args):
        stages.append("inspect")
        assert args == (source, "Data", "지역", "%_최신", tmp_path)
        return plan

    def execute(actual, overwrite, progress):
        stages.append("execute")
        assert actual is plan and not overwrite
        progress(1, 1, "서울")
        return SplitResult((target,), ())

    gui.controller._service = SimpleNamespace(preview=inspect, execute=execute, shutdown=lambda: None)
    gui.controller.state = replace(gui.controller.state, source=source, sheets=("Data",), sheet_name="Data",
                                   columns=("지역",), column_name="지역", output_dir=tmp_path)
    gui._render_state(gui.controller.state)
    gui.pattern_var.set("%_최신")
    gui.output_var.set(str(tmp_path / "missing"))
    monkeypatch.setattr("excel_splitter.gui.messagebox.askyesno", lambda *_, **__: False)
    gui.split_button.invoke()
    assert not gui._busy and not gui.split_button.instate(["disabled"])
    assert not (tmp_path / "missing").exists() and stages == []
    gui.output_var.set(str(tmp_path))
    monkeypatch.setattr("excel_splitter.gui.messagebox.askyesno", lambda *_, **__: stages.append("confirm") or True)
    assert not hasattr(gui, "preview_button") and not hasattr(gui, "merge_preview_button")
    gui.split_button.invoke()
    assert gui._busy
    complete_worker(gui)
    assert stages == ["inspect", "confirm", "execute"]
    assert gui.result_paths["split"] == (target,)


def test_execution_reconfigures_each_selected_progress_bar(toolkit):
    gui, _ = toolkit
    gui._progress_mode = "indeterminate"
    gui._progress_running = False
    gui.notebook.select(3)
    gui.etc_progress.configure(mode="determinate")

    gui._start_worker(lambda: ("noop", None), execution=True)

    assert str(gui.etc_progress["mode"]) == "indeterminate"
    assert gui.events.get(timeout=1) == ("ok", ("noop", None))
    gui._set_busy(False)
    assert str(gui.etc_progress["mode"]) == "determinate"


def test_one_click_merge_only_confirms_overwrite_and_rechecks_on_retry(toolkit, monkeypatch):
    gui, calls = toolkit
    add_files(gui, monkeypatch)
    original_preview = gui.merge_service.preview
    def preview(*args):
        plan = original_preview(*args)
        plan.prior_signature = object()
        return plan
    gui.merge_service.preview = preview
    prompts = []
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.askyesno", lambda _, text, **__: prompts.append(text) or False)
    gui._merge()
    complete_worker(gui)
    assert len(calls) == 1
    assert "덮어" in prompts[0]
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.askyesno", lambda *_, **__: True)
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda *_, **__: None)
    gui._merge()
    complete_worker(gui)
    assert len(calls) == 3 and calls[-1][1] is True
    assert gui.merge_preview is None
    assert str(gui.merge_button["state"]) == "normal"


def test_merge_preflight_failure_stays_visible_and_start_retries(toolkit, monkeypatch):
    from excel_splitter.errors import WorkbookValidationError

    gui, _ = toolkit
    add_files(gui, monkeypatch)
    original_preview = gui.merge_service.preview
    gui.merge_service.preview = lambda *_: (_ for _ in ()).throw(WorkbookValidationError("열 이름과 순서가 다릅니다"))
    gui.merge_button.invoke()
    complete_worker(gui)
    assert gui.merge_preview is None
    assert gui.notebook.tab(0, "state") == "normal"
    assert str(gui.merge_button["state"]) == "normal"
    assert "열 이름과 순서" in gui.error_detail
    gui.merge_service.preview = original_preview
    gui.merge_button.invoke()
    complete_worker(gui)
    assert gui.result_paths["merge"]


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
    summary = gui.result_panels["compare"][1]["text"]
    assert str(output) in summary and "3" in summary and "누락 시트" in summary
    assert messages == []
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
    summary = gui.result_panels["compare"][1]["text"]
    assert "3" in summary and "old_amount" in summary
    assert messages == []


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

    def execute(path, sheet_name, target, *, remove_artifacts, reset_fill, exclude_table_headers, remove_conditional_formats, progress):
        calls.append((path, sheet_name, target, remove_artifacts, reset_fill, exclude_table_headers, remove_conditional_formats))
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
    assert calls == [(source, "Clean", output, True, True, True, False)]
    assert str(output) in gui.result_panels["etc"][1]["text"]
    assert messages == []
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


def test_etc_conditional_format_option_defaults_runs_alone_and_locks_while_busy(toolkit, monkeypatch, tmp_path):
    gui, _ = toolkit
    gui.notebook.select(3)
    gui.etc_source_var.set(str(tmp_path / "source.xlsx"))
    gui.etc_output_var.set(str(tmp_path / "result.xlsx"))
    gui._handle_ok(("etc_source", ("Data",)))
    assert not gui.etc_remove_conditional_formats_var.get()
    assert gui.etc_remove_conditional_formats_checkbox.master is gui.etc_reset_fill_checkbox.master
    assert gui.etc_remove_conditional_formats_checkbox.grid_info()["row"] > gui.etc_reset_fill_checkbox.grid_info()["row"]
    assert str(gui.etc_button["state"]) == "disabled"
    calls = []

    def execute(source, sheet_name, target, **options):
        calls.append(options["remove_conditional_formats"])
        return target

    gui.etc_service = SimpleNamespace(execute=execute)
    monkeypatch.setattr("excel_splitter.toolkit_gui.messagebox.showinfo", lambda *_, **__: None)
    gui.etc_remove_conditional_formats_var.set(True)
    gui._render_etc_state()
    assert str(gui.etc_button["state"]) == "normal"
    gui._run_etc()
    assert str(gui.etc_remove_conditional_formats_checkbox["state"]) == "disabled"
    complete_worker(gui)
    assert calls == [True]
    assert not gui.etc_remove_conditional_formats_checkbox.instate(["disabled"])


def test_etc_table_header_option_is_enabled_only_when_reset_fill_is_checked(toolkit):
    gui, _ = toolkit
    header = gui.etc_exclude_table_headers_checkbox
    assert header.instate(["disabled"]) and gui.etc_exclude_table_headers_var.get()
    header.invoke()
    assert gui.etc_exclude_table_headers_var.get()

    gui.etc_reset_fill_checkbox.invoke()
    assert not header.instate(["disabled"])
    header.invoke()
    assert not gui.etc_exclude_table_headers_var.get()

    gui.etc_reset_fill_checkbox.invoke()
    assert header.instate(["disabled"])
    gui._set_busy(True)
    gui._set_busy(False)
    assert header.instate(["disabled"]) and not gui.etc_exclude_table_headers_var.get()
    gui.etc_reset_fill_checkbox.invoke()
    assert not header.instate(["disabled"]) and not gui.etc_exclude_table_headers_var.get()


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
    assert str(gui.split_button["state"]) == "disabled"

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
    gui.merge_button.invoke()
    complete_worker(gui)
    assert calls[-2][1] == merge_output


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


def test_position_mode_hides_keys_and_radio_switch_shows_them(toolkit):
    gui, _ = toolkit
    assert not gui.compare_key_options.winfo_manager()
    gui.compare_key_radio.invoke()
    assert gui.compare_by_key_var.get()
    assert gui.compare_key_options.winfo_manager() == "grid"
    gui.compare_position_radio.invoke()
    assert not gui.compare_by_key_var.get()
    assert not gui.compare_key_options.winfo_manager()


def test_output_validation_rejects_original_existing_and_missing_parent(toolkit, tmp_path):
    gui, _ = toolkit
    source = tmp_path / "source.xlsx"
    source.touch()
    gui.compare_reference_var.set(str(source))
    gui.compare_comparison_var.set(str(tmp_path / "other.xlsx"))
    for target in (source, tmp_path / "missing" / "result.xlsx", tmp_path / "wrong.csv"):
        gui.compare_output_var.set(str(target))
        assert gui.compare_button.instate(["disabled"])
        assert gui.compare_path_note.get()
    gui.compare_output_var.set(str(tmp_path / "fresh.xlsx"))
    assert not gui.compare_button.instate(["disabled"])


def test_completion_panel_keeps_result_and_merge_handoff_preserves_reference(toolkit, tmp_path, monkeypatch):
    gui, _ = toolkit
    target = tmp_path / "merged.xlsx"
    target.touch()
    reference = str(tmp_path / "reference.xlsx")
    gui.compare_reference_var.set(reference)
    gui.merge_output_var.set(str(target))
    gui._handle_ok(("merge_execute", target))
    assert gui.result_paths["merge"] == (target,)
    assert Path(gui.merge_output_var.get()) != target
    gui._use_merge_for_compare()
    assert gui.compare_reference_var.get() == reference
    assert gui.compare_comparison_var.get() == str(target)
    assert gui.notebook.index(gui.notebook.select()) == 2
    opened = []
    monkeypatch.setattr("excel_splitter.gui.os.startfile", opened.append)
    gui._open_result("merge")
    gui._open_result("merge", folder=True)
    assert opened == [target, target.parent]


def test_progress_unknown_phase_and_error_details_stay_readable(toolkit, monkeypatch):
    gui, _ = toolkit
    gui._executing = True
    gui._set_busy(True)
    gui._show_progress(0, 0, "파일 저장 중")
    assert str(gui.merge_progress["mode"]) == "indeterminate"
    assert "파일 저장 중" in gui.merge_status_var.get()
    gui._show_progress(2, 5, "데이터 병합 중")
    assert str(gui.merge_progress["mode"]) == "determinate"
    assert "2/5" in gui.merge_status_var.get()
    messages = []
    monkeypatch.setattr("excel_splitter.gui.messagebox.showerror", lambda _, text, **__: messages.append(text))
    raw = "COM diagnostic " * 1000
    gui._handle_error(RuntimeError(raw))
    assert messages == []
    assert len(gui.error_message_var.get()) < 1000
    assert raw in gui.error_detail
    assert gui.error_panel.winfo_manager() == "grid"
    assert not gui.error_text.winfo_manager()
    assert not gui._busy


def test_unknown_progress_moves_smoothly_without_restarting_on_phase_updates(toolkit):
    gui, _ = toolkit
    gui._executing = True
    gui._set_busy(True)
    bar = gui.merge_progress
    # Advance the real Tk animation by one frame, not a mocked progress widget.
    before = float(bar["value"])
    bar.step()
    assert 0 < (float(bar["value"]) - before) / float(bar["maximum"]) <= 0.05
    position = float(bar["value"])
    gui._show_progress(0, 0, "다음 파일 저장 중")
    assert float(bar["value"]) == position
    gui._set_busy(False)


def test_merge_progress_tracks_same_named_inputs_separately_and_only_animates_active_tab(toolkit, tmp_path):
    gui, _ = toolkit
    gui.merge_sources = [tmp_path / "서울" / "자료.xlsx", tmp_path / "부산" / "자료.xlsx"]
    gui._invalidate_merge_preview()
    gui._executing = True
    gui._set_busy(True)
    gui._show_progress(1, 2, "자료.xlsx")
    assert gui.merge_tree.set("0", "status") == "데이터 반영"
    assert gui.merge_tree.set("1", "status") == "병합 대기"
    gui._show_progress(2, 2, "자료.xlsx")
    assert gui.merge_tree.set("1", "status") == "데이터 반영"
    assert float(gui.merge_progress["value"]) == 2
    assert [float(bar["value"]) for bar in (gui.progress, gui.compare_progress, gui.etc_progress)] == [0, 0, 0]
    gui._set_busy(False)


def test_worker_status_and_errors_stay_on_the_active_tab(toolkit):
    gui, _ = toolkit
    split_status = gui.split_status_var.get()
    gui._executing = True
    gui._set_busy(True)
    gui._show_progress(1, 3, "병합 작업")
    assert gui.split_status_var.get() == split_status
    assert "병합 작업" in gui.merge_status_var.get()
    gui._handle_error(PermissionError("합성 잠금"))
    assert gui.split_status_var.get() == split_status


@pytest.mark.parametrize("tab", range(4))
@pytest.mark.parametrize("known_total", [False, True])
def test_real_worker_error_stops_animation_and_next_run_recovers(toolkit, tab, known_total):
    gui, _ = toolkit
    gui.notebook.select(tab)

    def fail_midway():
        gui.events.put(("progress", 2 if known_total else 0, 5 if known_total else 0, "중간 처리"))
        raise PermissionError("합성 파일 잠금 오류")

    gui._start_worker(fail_midway, execution=True)
    event = gui.events.get(timeout=3)
    gui._show_progress(*event[1:])
    event = gui.events.get(timeout=3)
    assert event[0] == "error"
    gui._handle_error(event[1])
    bar = gui._progress_widget()
    assert not gui._busy and not gui._executing and not gui._progress_running
    assert str(bar["mode"]) == "determinate" and float(bar["value"]) == 0
    done = tk.BooleanVar(value=False)
    gui.root.after(120, lambda: done.set(True))
    gui.root.wait_variable(done)
    assert float(bar["value"]) == 0
    assert all(gui.notebook.tab(index, "state") == "normal" for index in range(4))
    gui._show_progress(5, 5, "뒤늦은 알림")
    assert float(bar["value"]) == 0

    gui._start_worker(lambda: ("noop", None), execution=True)
    assert gui.events.get(timeout=3) == ("ok", ("noop", None))
    assert str(bar["mode"]) == "indeterminate"
    gui._show_progress(1, 1, "완료")
    gui._handle_ok(("noop", None))
    assert float(bar["value"]) == 1 and not gui._progress_running


def test_compare_details_keep_keys_and_both_coordinates_and_full_counts(toolkit, tmp_path):
    from excel_splitter.compare_service import CompareDifference, CompareResult

    gui, _ = toolkit
    gui.notebook.select(2)
    target = tmp_path / "result.xlsx"
    target.touch()
    gui.compare_by_key_var.set(True)
    detail = CompareDifference("changed", "Data", cell="E4", key="id=A-1", column_name="amount",
                               reference_value=10, comparison_value=20, reference_cell="B2", comparison_cell="E4")
    result = CompareResult(target, 5, (), missing_rows=2, added_rows=3, modified_cells=1,
                           details=(detail,), details_truncated=True, omitted_details=4)
    gui._handle_ok(("compare_execute", result))
    frame, label, tree, _ = gui.result_panels["compare"]
    text = label["text"]
    assert "값 변경 1셀" in text and "추가 3행" in text and "누락 2행" in text and "4건" in text
    row = tree.item(tree.get_children()[0], "values")
    assert all(value in row[1] for value in ("id=A-1", "B2", "E4"))
    assert Path(gui.compare_output_var.get()) != target
    assert "다음 실행" in gui.compare_path_note.get()


def test_position_summary_counts_highlighted_added_and_removed_cells(toolkit, tmp_path):
    from excel_splitter.compare_service import CompareResult

    gui, _ = toolkit
    gui._handle_ok(("compare_execute", CompareResult(tmp_path / "result.xlsx", changed_cells=7, missing_sheets=(), modified_cells=0)))
    summary = gui.result_panels["compare"][1]["text"]
    assert "7셀" in summary
    assert "추가 0행" not in summary


def test_error_details_include_underlying_cause(toolkit, monkeypatch):
    from excel_splitter.errors import WorkbookValidationError

    gui, _ = toolkit
    monkeypatch.setattr("excel_splitter.gui.messagebox.showerror", lambda *_, **__: None)
    try:
        try:
            raise RuntimeError("underlying COM diagnostic")
        except RuntimeError as exc:
            raise WorkbookValidationError("저장 실패") from exc
    except WorkbookValidationError as exc:
        gui._handle_error(exc)
    assert "underlying COM diagnostic" in gui.error_detail
    gui._copy_error()
    assert "underlying COM diagnostic" in gui.root.clipboard_get()


def test_split_failure_selection_never_opens_another_successful_file(toolkit, tmp_path, monkeypatch):
    from excel_splitter.models import SplitFailure, SplitResult

    gui, _ = toolkit
    path = tmp_path / "success.xlsx"
    gui._show_summary(SplitResult((path,), (SplitFailure("bad", "읽기 실패"),)))
    tree = gui.result_panels["split"][2]
    tree.selection_set(tree.get_children()[1])
    opened = []
    monkeypatch.setattr("excel_splitter.gui.os.startfile", opened.append)
    gui._open_result("split")
    assert opened == []
    gui._open_result("split", folder=True)
    assert opened == [tmp_path]
    tree.selection_set(tree.get_children()[0])
    gui._open_result("split")
    assert opened == [tmp_path, path]


def test_readonly_sources_show_filename_separately_from_long_path(toolkit, tmp_path):
    gui, _ = toolkit
    path = tmp_path / "long-parent-name" / "source-name.xlsx"
    gui.compare_reference_var.set(str(path))
    assert gui.source_name_labels[str(gui.compare_reference_var)]["text"] == "source-name.xlsx"


def _widgets(parent):
    for child in parent.winfo_children():
        yield child
        yield from _widgets(child)


def test_long_paths_are_wrapped_in_full_without_squeezing_input_controls(toolkit, tmp_path):
    from tkinter import ttk

    gui, _ = toolkit
    path = str(tmp_path / ("한글 경로 " * 6) / ("긴 합성 파일명 " * 5 + ".xlsx"))
    gui.compare_reference_var.set(path)
    gui.compare_comparison_var.set(path)
    gui.notebook.select(2)
    gui.root.geometry("900x680")
    gui.root.deiconify()
    gui.root.update()
    try:
        labels = [widget for widget in _widgets(gui.root) if isinstance(widget, ttk.Label)
                  and str(widget.cget("textvariable")) == str(gui.compare_reference_var)]
        assert labels, "Long paths need a full wrapped display beside the scrollable entry"
        label = labels[0]
        assert label.winfo_manager() == "grid"
        assert gui.root.getvar(str(label.cget("textvariable"))) == path
        assert float(label.cget("wraplength")) <= label.winfo_width()
        entries = [widget for widget in _widgets(gui.root) if isinstance(widget, ttk.Entry)
                   and str(widget.cget("textvariable")) == str(gui.compare_reference_var)]
        assert entries[0].winfo_width() >= 200
        assert entries[0].winfo_rootx() + entries[0].winfo_width() < gui.root.winfo_rootx() + gui.root.winfo_width()
    finally:
        gui.root.withdraw()


def test_merge_selection_exposes_full_path_in_compact_window(toolkit, tmp_path):
    from tkinter import ttk

    gui, _ = toolkit
    paths = [tmp_path / ("한글 폴더 " * 8) / f"합성 파일 {index}.xlsx" for index in range(2)]
    gui.merge_sources = paths
    gui._invalidate_merge_preview()
    gui.merge_tree.selection_set("1")
    gui.merge_tree.event_generate("<<TreeviewSelect>>")
    gui.root.geometry("900x680")
    gui.root.deiconify()
    gui.root.update()
    try:
        labels = [widget for widget in _widgets(gui.root) if isinstance(widget, ttk.Label)
                  and widget.cget("textvariable")
                  and gui.root.getvar(str(widget.cget("textvariable"))) == str(paths[1])]
        assert labels, "The selected merge input must be readable without truncation"
        assert float(labels[0].cget("wraplength")) <= labels[0].winfo_width()
        assert gui.merge_tree.cget("xscrollcommand")
    finally:
        gui.root.withdraw()


def test_result_detail_window_shows_full_unclipped_values(toolkit, tmp_path):
    gui, _ = toolkit
    content = "긴 비교값 " * 150
    gui._show_result("compare", (tmp_path / "result.xlsx",), "완료",
                     (("값 변경", "key=1 · 기준 A1, 대상 Z999", "내용", content, "after"),))
    tree = gui.result_panels["compare"][2]
    tree.selection_set(tree.get_children()[0])
    window = gui._show_result_detail("compare")
    try:
        text = next(widget for widget in window.winfo_children() if isinstance(widget, tk.Text))
        assert content in text.get("1.0", "end")
        assert "Z999" in text.get("1.0", "end")
        assert str(text["state"]) == "disabled"
    finally:
        window.destroy()


def test_merge_multiple_selection_and_delete_are_safe_during_work(toolkit, monkeypatch):
    gui, _ = toolkit
    add_files(gui, monkeypatch)
    gui.merge_tree.selection_set(("0", "1"))
    gui._set_busy(True)
    gui.merge_tree.event_generate("<Delete>")
    gui._remove_merge_file()
    gui._clear_merge_files()
    assert len(gui.merge_sources) == 2
    gui._set_busy(False)
    gui._remove_merge_file()
    assert gui.merge_sources == [] and not gui.merge_tree.get_children()
    assert gui.merge_button.instate(["disabled"])


@pytest.mark.parametrize("tab,kind", enumerate(("split", "merge", "compare", "etc")))
def test_primary_action_stays_visible_when_settings_scroll(toolkit, tab, kind):
    gui, _ = toolkit
    gui.notebook.select(tab)
    gui.compare_by_key_var.set(True)
    gui._compare_mode_changed()
    gui.root.geometry("760x540")
    gui.root.deiconify()
    gui.root.update()
    try:
        button = getattr(gui, kind + "_button")
        assert button.winfo_ismapped()
        assert button.winfo_rooty() + button.winfo_height() <= gui.root.winfo_rooty() + gui.root.winfo_height()
        assert button.winfo_rootx() + button.winfo_width() <= gui.root.winfo_rootx() + gui.root.winfo_width()
    finally:
        gui.root.withdraw()


def test_pasting_split_pattern_invalidates_preview_without_key_release(toolkit, tmp_path):
    gui, _ = toolkit
    gui.controller.state = replace(gui.controller.state, source=tmp_path / "source.xlsx", sheets=("Data",),
                                   sheet_name="Data", columns=("Team",), column_name="Team",
                                   output_dir=tmp_path, preview=object())
    gui._render_state(gui.controller.state)
    assert not gui.split_button.instate(["disabled"])
    gui.pattern_entry.delete(0, "end")
    gui.pattern_entry.insert(0, "%_새이름")
    assert gui.controller.state.pattern == "%_새이름"
    assert gui.controller.state.preview is None
    assert not gui.split_button.instate(["disabled"])


def test_callback_error_keeps_worker_locked_and_next_completion_recovers(toolkit):
    import threading
    from tkinter import ttk

    gui, _ = toolkit
    release = threading.Event()
    gui._start_worker(lambda: (release.wait(3), ("noop", None))[1], execution=True)
    button = ttk.Button(gui.root, command=lambda: (_ for _ in ()).throw(RuntimeError("callback failure")))
    try:
        button.invoke()
        assert "callback failure" in gui.error_detail
        assert gui.error_panel.winfo_manager() == "grid"
        assert gui._busy and gui.merge_button.instate(["disabled"])
    finally:
        release.set()
        button.destroy()
    complete_worker(gui)
    assert not gui._busy


def test_poll_recovers_after_result_rendering_failure(toolkit, monkeypatch):
    gui, _ = toolkit
    gui._start_worker(lambda: ("noop", None))
    event = gui.events.get(timeout=3)
    gui.events.put(event)
    monkeypatch.setattr(gui, "_handle_ok", lambda _: (_ for _ in ()).throw(RuntimeError("render failure")))
    gui.poll_queue()
    assert "render failure" in gui.error_detail and not gui._busy
    assert gui.root.tk.call("after", "info"), "Result rendering must not stop the event pump"
