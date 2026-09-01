import os
from pathlib import Path
import subprocess
import sys
from types import SimpleNamespace

import pytest

from excel_splitter.controller import AppController
from excel_splitter.gui import ExcelSplitterGui
from excel_splitter.parallel_writer import ParallelWriteAborted
from excel_splitter.models import (
    CanonicalKey,
    FileSignature,
    GroupSummary,
    OutputTarget,
    Preview,
    SplitResult,
    TableInfo,
    WorkbookSnapshot,
)


class _StateDouble:
    def __init__(self) -> None:
        self.state = "normal"

    def configure(self, *, state: str) -> None:
        self.state = state


class _RootDouble:
    def protocol(self, *_args) -> None:
        pass

    def destroy(self) -> None:
        pass


class _VariableDouble:
    def __init__(self, value: object = None) -> None:
        self.value = value

    def set(self, value: object) -> None:
        self.value = value


class _ProgressDouble:
    def __init__(self, maximum: int = 1) -> None:
        self.maximum = maximum

    def configure(self, *, maximum: int) -> None:
        self.maximum = maximum


def test_leaving_busy_state_restores_browse_buttons() -> None:
    gui = object.__new__(ExcelSplitterGui)
    gui.source_button = _StateDouble()
    gui.output_button = _StateDouble()
    gui._input_widgets = [gui.source_button, gui.output_button]
    gui.root = _RootDouble()
    gui.status_var = _VariableDouble()
    gui.controller = SimpleNamespace(state=object())
    gui._render_state = lambda _state: None

    gui._set_busy(True)
    gui._set_busy(False)

    assert gui.source_button.state == "normal"
    assert gui.output_button.state == "normal"


def test_idle_window_close_shuts_down_session_before_destroy() -> None:
    events: list[str] = []
    gui = object.__new__(ExcelSplitterGui)
    gui._busy = False
    gui.controller = SimpleNamespace(shutdown=lambda: events.append("shutdown"))
    gui.root = SimpleNamespace(destroy=lambda: events.append("destroy"))
    gui.logger = SimpleNamespace(exception=lambda *_args, **_kwargs: None)

    gui._on_close()

    assert events == ["shutdown", "destroy"]


def test_parallel_abort_error_reports_partial_and_unstarted_counts(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shown: list[str] = []
    gui = object.__new__(ExcelSplitterGui)
    gui._set_busy = lambda _busy: None
    gui.logger = SimpleNamespace(error=lambda *_args, **_kwargs: None)
    gui.root = object()
    gui.controller = SimpleNamespace(state=object())
    gui._render_state = lambda _state: None
    gui.progress_var = _VariableDouble()
    gui.progress = _ProgressDouble()
    gui.status_var = SimpleNamespace(set=lambda _value: None)
    partial = SplitResult((Path("done.xlsx"),), ())
    error = ParallelWriteAborted("worker failed", partial, (Path("later.xlsx"),))
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.showerror",
        lambda _title, message, **_kwargs: shown.append(message),
    )

    gui._handle_error(error)

    assert "완료 1개" in shown[0]
    assert "실패 0개" in shown[0]
    assert "시작하지 못함 1개" in shown[0]


@pytest.mark.parametrize(
    "error",
    [
        RuntimeError("unexpected"),
        ParallelWriteAborted(
            "worker failed",
            SplitResult((Path("done.xlsx"),), ()),
            (Path("later.xlsx"),),
        ),
    ],
    ids=["general", "parallel-abort"],
)
def test_error_resets_progress_and_next_progress_update_still_works(
    error: Exception,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    gui = object.__new__(ExcelSplitterGui)
    gui._set_busy = lambda _busy: None
    gui.logger = SimpleNamespace(error=lambda *_args, **_kwargs: None)
    gui.root = object()
    gui.controller = SimpleNamespace(state=object())
    gui._render_state = lambda _state: None
    gui.progress_var = _VariableDouble()
    gui.progress = _ProgressDouble()
    gui.status_var = _VariableDouble()

    def show_error(*_args: object, **_kwargs: object) -> None:
        assert gui.progress_var.value == 0
        assert gui.progress.maximum == 1

    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.showerror",
        show_error,
    )
    gui._show_progress(4, 5, "A")

    gui._handle_error(error)

    assert gui.progress_var.value == 0
    assert gui.progress.maximum == 1

    gui._show_progress(1, 3, "B")
    assert gui.progress_var.value == 1
    assert gui.progress.maximum == 3


def test_script_entry_calls_main_without_requiring_package_context() -> None:
    project_root = Path(__file__).parents[2]
    environment = os.environ.copy()
    environment["PYTHONPATH"] = str(project_root / "src")
    probe = "\n".join(
        (
            "import runpy, sys, types",
            "import excel_splitter",
            "calls = []",
            "stub = types.ModuleType('excel_splitter.app')",
            "stub.main = lambda: calls.append('called')",
            "sys.modules['excel_splitter.app'] = stub",
            "excel_splitter.app = stub",
            "runpy.run_path('src/excel_splitter/__main__.py', run_name='__main__')",
            "assert calls == ['called'], calls",
        )
    )

    completed = subprocess.run(
        [sys.executable, "-c", probe],
        cwd=project_root,
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    assert "attempted relative import" not in completed.stderr


class FakeService:
    def __init__(self) -> None:
        self.preview_value = _preview()
        self.execute_calls: list[tuple[Preview, bool]] = []

    def list_sheets(self, source: Path) -> tuple[str, ...]:
        return ("분류표", "참조")

    def inspect_sheet(self, source: Path, sheet_name: str) -> TableInfo:
        return TableInfo(sheet_name, "Table1", ("구분", "금액"), 2)

    def preview(
        self,
        source: Path,
        sheet_name: str,
        column_name: str,
        pattern: str,
        output_dir: Path,
    ) -> Preview:
        return self.preview_value

    def execute(self, preview, overwrite, progress):
        self.execute_calls.append((preview, overwrite))
        progress(1, 1, "A")
        return SplitResult((Path("A_분할.xlsx"),), ())

    def shutdown(self):
        self.shutdown_called = True


def _preview() -> Preview:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source=Path("a.xlsx"),
        signature=FileSignature(1, 2, "abc"),
        sheet_name="분류표",
        table_name="Table1",
        column_name="구분",
        row_count=2,
        groups=(GroupSummary(key, "A", 2, (1, 2)),),
    )
    return Preview(
        snapshot=snapshot,
        targets=(OutputTarget(key, "A", Path("A_분할.xlsx"), None),),
        collisions=(),
    )


def _ready_controller(service: FakeService) -> AppController:
    controller = AppController(service)
    controller.select_source(Path("folder/a.xlsx"))
    controller.select_sheet("분류표")
    controller.select_column("구분")
    return controller


def test_selecting_new_file_resets_downstream_state() -> None:
    controller = AppController(FakeService())
    controller.select_source(Path("a.xlsx"))
    controller.select_sheet("분류표")
    controller.select_source(Path("b.xlsx"))
    assert controller.state.sheet_name is None
    assert controller.state.columns == ()
    assert controller.state.preview is None


def test_selecting_source_loads_sheets_and_defaults_output_folder() -> None:
    controller = AppController(FakeService())

    controller.select_source(Path("folder/a.xlsx"))

    assert controller.state.source == Path("folder/a.xlsx")
    assert controller.state.sheets == ("분류표", "참조")
    assert controller.state.output_dir == Path("folder")


def test_selecting_sheet_loads_columns_and_resets_column_and_preview() -> None:
    service = FakeService()
    controller = _ready_controller(service)
    controller.create_preview()

    controller.select_sheet("참조")

    assert controller.state.columns == ("구분", "금액")
    assert controller.state.column_name is None
    assert controller.state.preview is None


@pytest.mark.parametrize("change", ["column", "output", "pattern"])
def test_upstream_edit_invalidates_preview(change: str) -> None:
    service = FakeService()
    controller = _ready_controller(service)
    controller.create_preview()

    if change == "column":
        controller.select_column("금액")
    elif change == "output":
        controller.select_output_dir(Path("other"))
    else:
        controller.set_pattern("%_결과")

    assert controller.state.preview is None


def test_create_preview_stores_service_result() -> None:
    service = FakeService()
    controller = _ready_controller(service)

    result = controller.create_preview()

    assert result is service.preview_value
    assert controller.state.preview is service.preview_value


def test_execute_uses_current_preview_and_forwards_progress() -> None:
    service = FakeService()
    controller = _ready_controller(service)
    current = controller.create_preview()
    progress_events: list[tuple[int, int, str]] = []

    result = controller.execute(True, lambda *event: progress_events.append(event))

    assert service.execute_calls == [(current, True)]
    assert progress_events == [(1, 1, "A")]
    assert result.succeeded == (Path("A_분할.xlsx"),)


def test_execute_rejects_missing_or_invalidated_preview() -> None:
    service = FakeService()
    controller = _ready_controller(service)

    with pytest.raises(RuntimeError, match="미리보기"):
        controller.execute(False, lambda *_: None)

    controller.create_preview()
    controller.set_pattern("%_결과")
    with pytest.raises(RuntimeError, match="미리보기"):
        controller.execute(False, lambda *_: None)


def test_controller_shutdown_delegates_to_service() -> None:
    service = FakeService()
    controller = AppController(service)
    controller.shutdown()
    assert service.shutdown_called is True
