from pathlib import Path

import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.file_signature import capture_signature
from excel_splitter.models import (
    CanonicalKey,
    GroupSummary,
    SplitResult,
    TableInfo,
    WorkbookSnapshot,
)
from excel_splitter.split_service import SplitService


class FakeGateway:
    def __init__(self, snapshot: WorkbookSnapshot) -> None:
        self.snapshot = snapshot
        self.written = False

    def list_worksheets(self, source: Path) -> tuple[str, ...]:
        return ("분류표",)

    def inspect_table(self, source: Path, sheet_name: str) -> TableInfo:
        return TableInfo(sheet_name, "Table1", ("구분",), 1)

    def build_snapshot(
        self, source: Path, sheet_name: str, column_name: str
    ) -> WorkbookSnapshot:
        return self.snapshot

    def write_groups(self, snapshot, targets, progress):
        self.written = True
        return SplitResult(tuple(target.path for target in targets), ())


def _service(tmp_path: Path) -> tuple[Path, FakeGateway, SplitService]:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    signature = capture_signature(source)
    group = GroupSummary(CanonicalKey("text", "A"), "A", 1, (1,))
    snapshot = WorkbookSnapshot(
        source, signature, "분류표", "Table1", "구분", 1, (group,)
    )
    gateway = FakeGateway(snapshot)
    return source, gateway, SplitService(gateway)


def test_execute_requires_overwrite_approval_when_collision_exists(
    tmp_path: Path,
) -> None:
    source, gateway, service = _service(tmp_path)
    existing = tmp_path / "A.xlsx"
    existing.write_bytes(b"existing")

    preview = service.preview(source, "분류표", "구분", "%", tmp_path)

    assert preview.targets[0].prior_signature == capture_signature(existing)
    assert preview.collisions == (existing,)
    with pytest.raises(WorkbookValidationError, match="덮어쓰기 승인"):
        service.execute(preview, overwrite=False, progress=lambda *_: None)
    assert not gateway.written


def test_execute_rejects_source_changed_after_preview_before_writing(
    tmp_path: Path,
) -> None:
    source, gateway, service = _service(tmp_path)
    preview = service.preview(source, "분류표", "구분", "%", tmp_path)
    source.write_bytes(source.read_bytes() + b"!")

    with pytest.raises(WorkbookValidationError, match="미리보기 이후 변경"):
        service.execute(preview, overwrite=True, progress=lambda *_: None)

    assert not gateway.written


def test_preview_rejects_missing_output_directory_before_building_targets(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source, _gateway, service = _service(tmp_path)
    monkeypatch.setattr(
        "excel_splitter.split_service.build_targets",
        lambda *_args: pytest.fail("targets must not be built"),
    )

    with pytest.raises(WorkbookValidationError, match="출력 폴더"):
        service.preview(
            source,
            "분류표",
            "구분",
            "%",
            tmp_path / "missing",
        )


def test_main_builds_default_service_and_runs_gui(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from excel_splitter import app

    events: list[object] = []
    gateway = object()
    service = object()
    controller = object()
    root = type("Root", (), {"mainloop": lambda self: events.append("mainloop")})()
    monkeypatch.setattr(app.tk, "Tk", lambda: root)
    monkeypatch.setattr(app, "ExcelComGateway", lambda: gateway)
    monkeypatch.setattr(
        app,
        "SplitService",
        lambda actual_gateway: events.append(actual_gateway) or service,
    )
    monkeypatch.setattr(
        app,
        "AppController",
        lambda actual_service: events.append(actual_service) or controller,
    )
    monkeypatch.setattr(
        app,
        "ExcelSplitterGui",
        lambda actual_root, actual_controller: events.append(
            (actual_root, actual_controller)
        ),
    )

    app.main()

    assert events == [gateway, service, (root, controller), "mainloop"]
