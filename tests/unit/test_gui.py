from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_splitter.gui import DELETED_SHEETS_WARNING, ExcelSplitterGui
from excel_splitter.errors import WorkbookValidationError
from excel_splitter.models import FileSignature, Preview, WorkbookSnapshot


@pytest.mark.parametrize("has_artifacts", (False, True))
def test_split_confirmation_warns_only_when_artifacts_will_be_deleted(
    has_artifacts: bool, monkeypatch: pytest.MonkeyPatch
) -> None:
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Data",
        "Orders",
        "Team",
        0,
        (),
        has_removable_artifacts=has_artifacts,
    )
    preview = Preview(snapshot, (), ())
    gui = ExcelSplitterGui.__new__(ExcelSplitterGui)
    gui.controller = SimpleNamespace(state=SimpleNamespace(preview=preview))
    gui.root = object()
    prompts: list[str] = []
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno",
        lambda _title, prompt, **_options: prompts.append(prompt) or False,
    )

    gui._split()

    expected = DELETED_SHEETS_WARNING
    if has_artifacts:
        expected += "\n\n메모 및 도형은 삭제됩니다"
    assert prompts == [expected]


def _preview_gui(output_dir: Path, workers: list[object]) -> ExcelSplitterGui:
    gui = ExcelSplitterGui.__new__(ExcelSplitterGui)
    gui.root = object()
    gui.controller = SimpleNamespace(
        state=SimpleNamespace(output_dir=output_dir),
        set_pattern=lambda _pattern: None,
        create_preview=lambda: "preview",
    )
    gui.pattern_var = SimpleNamespace(get=lambda: "%_분할")
    gui._clear_preview = lambda: None
    gui._start_worker = workers.append
    return gui


def test_preview_creates_missing_nested_output_folder_before_starting_worker(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output_dir = tmp_path / "new" / "nested" / "output"
    workers: list[object] = []
    gui = _preview_gui(output_dir, workers)
    started: list[object] = []

    def start_worker(action: object) -> None:
        assert output_dir.is_dir()
        started.append(action())

    gui._start_worker = start_worker
    prompts: list[str] = []
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno",
        lambda _title, prompt, **_options: prompts.append(prompt) or True,
    )

    gui._preview()

    assert output_dir.is_dir()
    assert started == [("preview", "preview")]
    assert prompts == [f"폴더가 존재하지 않습니다. 만드시겠습니까?\n\n{output_dir}"]


def test_preview_does_not_create_folder_or_start_worker_when_creation_is_declined(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output_dir = tmp_path / "missing"
    workers: list[object] = []
    gui = _preview_gui(output_dir, workers)
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno", lambda *_args, **_options: False
    )

    gui._preview()

    assert not output_dir.exists()
    assert workers == []


def test_preview_does_not_prompt_for_existing_output_folder(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workers: list[object] = []
    gui = _preview_gui(tmp_path, workers)
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno",
        lambda *_args, **_options: pytest.fail("existing folder prompted"),
    )

    gui._preview()

    assert len(workers) == 1


def test_preview_reports_folder_creation_failure_without_starting_worker(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output_dir = tmp_path / "missing"
    workers: list[object] = []
    gui = _preview_gui(output_dir, workers)
    errors: list[object] = []
    gui._handle_error = errors.append
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno", lambda *_args, **_options: True
    )

    def fail_mkdir(self: Path, *_args: object, **_kwargs: object) -> None:
        raise OSError("access denied")

    monkeypatch.setattr(Path, "mkdir", fail_mkdir)

    gui._preview()

    assert workers == []
    assert len(errors) == 1
    assert isinstance(errors[0], WorkbookValidationError)
    assert str(output_dir) in str(errors[0]) and "access denied" in str(errors[0])


def test_preview_keeps_existing_file_for_backend_output_validation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output_path = tmp_path / "not-a-folder"
    output_path.write_text("file")
    workers: list[object] = []
    gui = _preview_gui(output_path, workers)
    gui.controller.create_preview = lambda: (_ for _ in ()).throw(
        WorkbookValidationError("출력 폴더가 존재하지 않습니다.")
    )
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno",
        lambda *_args, **_options: pytest.fail("existing file prompted"),
    )

    gui._preview()

    assert len(workers) == 1
    with pytest.raises(WorkbookValidationError, match="출력 폴더가 존재하지 않습니다."):
        workers[0]()
