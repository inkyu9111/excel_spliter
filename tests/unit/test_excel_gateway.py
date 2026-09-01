from __future__ import annotations

import sys
import os
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace

import pytest

import excel_splitter.excel_gateway as excel_gateway

from excel_splitter.errors import (
    SplitExecutionError,
    WorkbookValidationError,
)
from excel_splitter.file_signature import capture_signature
from excel_splitter.models import (
    CanonicalKey,
    FileSignature,
    GroupSummary,
    OutputTarget,
    SplitResult,
    TableInfo,
    WorkbookSnapshot,
)
from excel_splitter.excel_gateway import (
    ExcelComGateway,
    _column_index,
    _copy_to_master,
    _delete_rows,
    _excel_error_code,
    _excel_session,
    _open_workbook,
    _publish_temp,
    _remove_other_sheets,
    _restore_shapes,
    _shape_snapshot,
    _single_table,
    _validate_below_table,
    _verify_target_unchanged,
    _write_one_group,
)


def test_validate_below_table_uses_bulk_properties_for_an_empty_range() -> None:
    calls: list[str] = []

    class Bounded:
        @property
        def Value2(self): calls.append("Value2"); return ((None, None),)
        @property
        def Formula(self): calls.append("Formula"); return ((None, None),)
        @property
        def MergeCells(self): calls.append("MergeCells"); return False
        @property
        def Hyperlinks(self): calls.append("Hyperlinks"); return SimpleNamespace(Count=0)
        @property
        def Style(self): calls.append("Style"); return "Normal"

    bounded = Bounded()
    sheet = SimpleNamespace(
        UsedRange=SimpleNamespace(Row=1, Rows=SimpleNamespace(Count=3)),
        Cells=lambda *_args: object(), Range=lambda *_args: bounded,
        Comments=SimpleNamespace(Count=0),
    )
    table = SimpleNamespace(Range=SimpleNamespace(
        Row=1, Column=1, Rows=SimpleNamespace(Count=1),
        Columns=SimpleNamespace(Count=2),
    ))

    _validate_below_table(sheet, table)

    assert calls == ["Value2", "Formula", "MergeCells", "Hyperlinks", "Style"]


def test_validate_below_table_ignores_unsupported_threaded_comment_model() -> None:
    com_error = type("com_error", (Exception,), {})

    class LegacySheet:
        UsedRange = SimpleNamespace(Row=1, Rows=SimpleNamespace(Count=2))
        Comments = SimpleNamespace(Count=0)
        Cells = staticmethod(lambda *_args: object())
        Range = staticmethod(lambda *_args: SimpleNamespace(
            Value2=None, Formula=None, MergeCells=False,
            Hyperlinks=SimpleNamespace(Count=0), Style="Normal",
        ))

        @property
        def CommentsThreaded(self):
            error = com_error(-2147352567, "스레드 주석 VBA 모델을 지원하지 않습니다")
            error.hresult = -2147352567
            raise error

    table = SimpleNamespace(Range=SimpleNamespace(
        Row=1, Column=1, Rows=SimpleNamespace(Count=1),
        Columns=SimpleNamespace(Count=1),
    ))

    _validate_below_table(LegacySheet(), table)


def test_validate_below_table_propagates_unexpected_threaded_comment_error() -> None:
    bounded = SimpleNamespace(
        Value2=None, Formula=None, MergeCells=False,
        Hyperlinks=SimpleNamespace(Count=0), Style="Normal",
    )
    error = RuntimeError("COM transport lost")

    class BrokenSheet:
        UsedRange = SimpleNamespace(Row=1, Rows=SimpleNamespace(Count=2))
        Comments = SimpleNamespace(Count=0)
        Cells = staticmethod(lambda *_args: object())
        Range = staticmethod(lambda *_args: bounded)

        @property
        def CommentsThreaded(self):
            raise error

    sheet = BrokenSheet()
    table = SimpleNamespace(Range=SimpleNamespace(
        Row=1, Column=1, Rows=SimpleNamespace(Count=1),
        Columns=SimpleNamespace(Count=1),
    ))

    with pytest.raises(RuntimeError) as caught:
        _validate_below_table(sheet, table)
    assert caught.value is error


class Recorder:
    def __init__(self) -> None:
        self.deleted: list[int] = []

    def __call__(self, index: int):
        owner = self

        class Row:
            def Delete(self) -> None:
                owner.deleted.append(index)

        return Row()


def test_delete_rows_uses_descending_indexes() -> None:
    rows = Recorder()
    _delete_rows(rows, (1, 4, 2))
    assert rows.deleted == [4, 2, 1]


def test_open_workbook_disables_password_prompts() -> None:
    calls: list[tuple[str, dict[str, object]]] = []
    workbooks = SimpleNamespace(
        Open=lambda source, **options: calls.append((source, options)) or object()
    )

    _open_workbook(SimpleNamespace(Workbooks=workbooks), Path("source.xlsx"), read_only=True)

    assert calls[0][1]["Password"] == ""
    assert calls[0][1]["WriteResPassword"] == ""


def test_gateway_starts_injected_source_session_lazily() -> None:
    events: list[object] = []

    class Session:
        def start(self):
            events.append("start")

        def open_source(self, source):
            events.append(("open", source))
            return SimpleNamespace(sheets=())

        def shutdown(self):
            events.append("shutdown")

    gateway = ExcelComGateway(Session())
    assert events == []
    assert gateway.list_worksheets(Path("source.xlsx")) == ()
    assert events == ["start", ("open", Path("source.xlsx"))]
    gateway.shutdown()
    assert events[-1] == "shutdown"


def test_gateway_prewarm_starts_persistent_session_without_opening_source() -> None:
    events: list[str] = []

    class Session:
        def start(self): events.append("start")
        def shutdown(self): events.append("shutdown")

    gateway = ExcelComGateway(Session())
    gateway.prewarm()
    gateway.prewarm()
    gateway.shutdown()

    assert events == ["start", "shutdown"]


def test_failed_prewarm_is_retried_by_real_source_action() -> None:
    attempts: list[str] = []

    class Session:
        def start(self):
            attempts.append("start")
            raise SplitExecutionError("Excel unavailable")

        def shutdown(self): pass

    gateway = ExcelComGateway(Session())
    with pytest.raises(SplitExecutionError, match="Excel unavailable"):
        gateway.prewarm()
    with pytest.raises(SplitExecutionError, match="Excel unavailable"):
        gateway.list_worksheets(Path("source.xlsx"))

    assert attempts == ["start", "start"]


def test_excel_error_code_detaches_the_positive_cverr_code() -> None:
    assert _excel_error_code(-2146826246) == 2042
    assert _excel_error_code(2042) == 2042
    assert _excel_error_code(-2146826243) == 2045
    assert _excel_error_code(2045) == 2045
    assert _excel_error_code(1) is None


@pytest.mark.parametrize(
    ("count", "message"),
    [(0, "정식 Excel Table이 없습니다"), (2, "2개 이상")],
)
def test_single_table_rejects_invalid_table_count(count: int, message: str) -> None:
    tables = SimpleNamespace(Count=count)
    with pytest.raises(WorkbookValidationError, match=message):
        _single_table(tables)


def test_column_index_requires_one_exact_case_sensitive_match() -> None:
    columns = SimpleNamespace(
        Count=3,
        Item=lambda index: SimpleNamespace(Name=("Team", "team", "Team")[index - 1]),
    )
    table = SimpleNamespace(ListColumns=columns)

    assert _column_index(table, "team") == 2
    with pytest.raises(WorkbookValidationError, match="정확히 하나"):
        _column_index(table, "Team")
    with pytest.raises(WorkbookValidationError, match="정확히 하나"):
        _column_index(table, "TEAM")


def test_remove_other_sheets_deletes_from_end_and_keeps_selected() -> None:
    deleted: list[str] = []

    class Sheet:
        def __init__(self, name: str) -> None:
            self.Name = name

        def Delete(self) -> None:
            deleted.append(self.Name)

    sheets = [Sheet("first"), Sheet("keep"), Sheet("chart")]
    collection = SimpleNamespace(
        Count=len(sheets), Item=lambda index: sheets[index - 1]
    )

    _remove_other_sheets(collection, "keep")
    assert deleted == ["chart", "first"]


def test_excel_session_always_quits_and_uninitializes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    class Excel:
        def Quit(self) -> None:
            events.append("quit")

    pythoncom = SimpleNamespace(
        CoInitialize=lambda: events.append("initialize"),
        CoUninitialize=lambda: events.append("uninitialize"),
    )
    client = SimpleNamespace(DispatchEx=lambda _name: Excel())
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    with pytest.raises(SplitExecutionError, match="Excel 자동화"):
        with _excel_session():
            raise RuntimeError("open failed")

    assert events == ["initialize", "quit", "uninitialize"]


def test_excel_session_does_not_swallow_attribute_error_from_with_body(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    class Excel:
        def Quit(self) -> None:
            events.append("quit")

    pythoncom = SimpleNamespace(
        CoInitialize=lambda: events.append("initialize"),
        CoUninitialize=lambda: events.append("uninitialize"),
    )
    client = SimpleNamespace(DispatchEx=lambda _name: Excel())
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    with pytest.raises(SplitExecutionError, match="Excel 자동화"):
        with _excel_session():
            raise AttributeError("missing COM member")

    assert events == ["initialize", "quit", "uninitialize"]


def test_output_excel_session_uses_manual_calculation_without_changing_default(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created: list[SimpleNamespace] = []

    def dispatch(_name: str) -> SimpleNamespace:
        excel = SimpleNamespace(
            Calculation=-4105,
            CalculateBeforeSave=False,
            Quit=lambda: None,
        )
        created.append(excel)
        return excel

    pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    client = SimpleNamespace(DispatchEx=dispatch)
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    output_session = getattr(excel_gateway, "_output_excel_session", None)
    assert output_session is not None
    with output_session() as output_excel:
        assert output_excel.Calculation == -4135
        assert output_excel.CalculateBeforeSave is True
    with _excel_session() as source_excel:
        assert source_excel.Calculation == -4105
        assert source_excel.CalculateBeforeSave is False


def test_shape_snapshot_frees_shapes_before_rows_move_and_restores_all_properties() -> None:
    shape = SimpleNamespace(
        Name="Logo",
        Left=10.0,
        Top=20.0,
        Width=30.0,
        Height=40.0,
        Placement=1,
    )
    shapes = SimpleNamespace(Count=1, Item=lambda _key: shape)
    sheet = SimpleNamespace(Shapes=shapes)

    snapshot = _shape_snapshot(sheet)

    assert shape.Placement == 3
    shape.Left, shape.Top, shape.Width, shape.Height, shape.Placement = (
        1,
        2,
        3,
        4,
        2,
    )
    _restore_shapes(sheet, snapshot)
    assert (
        shape.Left,
        shape.Top,
        shape.Width,
        shape.Height,
        shape.Placement,
    ) == (10.0, 20.0, 30.0, 40.0, 1)


def test_restore_shapes_reports_missing_shape_name_and_stage() -> None:
    shape = SimpleNamespace(
        Name="Delete Button",
        Left=10.0,
        Top=20.0,
        Width=30.0,
        Height=40.0,
        Placement=1,
    )
    available: dict[object, object] = {1: shape, "Delete Button": shape}
    sheet = SimpleNamespace(
        Shapes=SimpleNamespace(Count=1, Item=lambda key: available[key])
    )
    snapshot = _shape_snapshot(sheet)
    del available["Delete Button"]

    with pytest.raises(SplitExecutionError, match="도형 복원.*Delete Button"):
        _restore_shapes(sheet, snapshot)


@pytest.mark.parametrize(
    ("failed_stage", "expected_message"),
    (
        ("open", "파일 열기"),
        ("delete", "행 삭제"),
        ("restore", "도형 복원"),
        ("save", "저장"),
    ),
)
def test_write_one_group_identifies_the_failed_com_stage(
    failed_stage: str,
    expected_message: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Sheet1",
        "Table1",
        "Team",
        1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", Path("result.xlsx"), None)
    sheet = SimpleNamespace(
        Shapes=SimpleNamespace(Count=0),
        FilterMode=False,
    )
    rows = SimpleNamespace(Count=1)
    table = SimpleNamespace(Name="Table1", ListRows=rows)

    class Workbook:
        Sheets = SimpleNamespace(Count=1)

        def Save(self) -> None:
            if failed_stage == "save":
                raise RuntimeError("COM save failed")

        def Close(self, **_kwargs) -> None:
            pass

    workbook = Workbook()
    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", lambda *_: None)
    if failed_stage == "open":
        monkeypatch.setattr(
            "excel_splitter.excel_gateway._open_workbook",
            lambda *_args, **_kwargs: (_ for _ in ()).throw(
                RuntimeError("COM open failed")
            ),
        )
    else:
        monkeypatch.setattr(
            "excel_splitter.excel_gateway._open_workbook",
            lambda *_args, **_kwargs: workbook,
        )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda *_args: (sheet, table),
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._publish_temp", lambda *_: None)
    if failed_stage == "delete":
        monkeypatch.setattr(
            "excel_splitter.excel_gateway._delete_rows",
            lambda *_: (_ for _ in ()).throw(RuntimeError("COM delete failed")),
        )
    if failed_stage == "restore":
        monkeypatch.setattr(
            "excel_splitter.excel_gateway._restore_shapes",
            lambda *_: (_ for _ in ()).throw(RuntimeError("COM shape failed")),
        )

    with pytest.raises(SplitExecutionError, match=expected_message):
        _write_one_group(SimpleNamespace(), Path("master.xlsx"), snapshot, target)


def _stub_successful_group_write(
    monkeypatch: pytest.MonkeyPatch,
    open_workbook,
    shape_snapshot=lambda _sheet: (),
) -> tuple[WorkbookSnapshot, OutputTarget]:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Sheet1",
        "Table1",
        "Team",
        1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", Path("result.xlsx"), None)
    sheet = SimpleNamespace(Shapes=SimpleNamespace(Count=0), FilterMode=False)
    table = SimpleNamespace(Name="Table1", ListRows=SimpleNamespace(Count=1))
    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", lambda *_: None)
    monkeypatch.setattr("excel_splitter.excel_gateway._open_workbook", open_workbook)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda *_args: (sheet, table),
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._shape_snapshot", shape_snapshot
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._restore_shapes", lambda *_: None)
    monkeypatch.setattr("excel_splitter.excel_gateway._publish_temp", lambda *_: None)
    return snapshot, target


def test_write_one_group_reapplies_manual_calculation_after_workbook_open(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    excel = SimpleNamespace(Calculation=-4135, CalculateBeforeSave=True)
    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: None,
        Close=lambda **_kwargs: None,
    )
    observed_before_delete: list[tuple[int, bool]] = []

    def open_workbook(*_args, **_kwargs):
        excel.Calculation = -4105
        excel.CalculateBeforeSave = False
        return workbook

    def shape_snapshot(_sheet):
        observed_before_delete.append(
            (excel.Calculation, excel.CalculateBeforeSave)
        )
        return ()

    snapshot, target = _stub_successful_group_write(
        monkeypatch, open_workbook, shape_snapshot
    )

    _write_one_group(excel, Path("master.xlsx"), snapshot, target)

    assert observed_before_delete == [(-4135, True)]


def test_write_one_group_rejects_calculation_settings_not_applied_after_open(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class Excel:
        def __init__(self) -> None:
            self._calculation = -4135
            self._calculate_before_save = True
            self.accept_settings = True

        @property
        def Calculation(self):
            return self._calculation

        @Calculation.setter
        def Calculation(self, value) -> None:
            if self.accept_settings:
                self._calculation = value

        @property
        def CalculateBeforeSave(self):
            return self._calculate_before_save

        @CalculateBeforeSave.setter
        def CalculateBeforeSave(self, value) -> None:
            if self.accept_settings:
                self._calculate_before_save = value

    excel = Excel()
    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: None,
        Close=lambda **_kwargs: None,
    )

    def open_workbook(*_args, **_kwargs):
        excel.Calculation = -4105
        excel.CalculateBeforeSave = False
        excel.accept_settings = False
        return workbook

    snapshot, target = _stub_successful_group_write(monkeypatch, open_workbook)

    with pytest.raises(SplitExecutionError, match="파일 열기"):
        _write_one_group(excel, Path("master.xlsx"), snapshot, target)


def test_parallel_group_saves_are_serialized(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Sheet1",
        "Table1",
        "Team",
        1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    targets = (
        OutputTarget(key, "A", Path("first.xlsx"), None),
        OutputTarget(key, "A", Path("second.xlsx"), None),
    )
    state_lock = threading.Lock()
    active_saves = 0
    maximum_active_saves = 0

    class Workbook:
        Sheets = SimpleNamespace(Count=1)

        def Save(self) -> None:
            nonlocal active_saves, maximum_active_saves
            with state_lock:
                active_saves += 1
                maximum_active_saves = max(maximum_active_saves, active_saves)
            time.sleep(0.05)
            with state_lock:
                active_saves -= 1

        def Close(self, **_kwargs) -> None:
            pass

    sheet = SimpleNamespace(Shapes=SimpleNamespace(Count=0), FilterMode=False)
    table = SimpleNamespace(Name="Table1", ListRows=SimpleNamespace(Count=1))
    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", lambda *_: None)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._open_workbook",
        lambda *_args, **_kwargs: Workbook(),
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda *_args: (sheet, table),
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._publish_temp", lambda *_: None)
    errors: list[Exception] = []

    def write(target: OutputTarget) -> None:
        try:
            _write_one_group(SimpleNamespace(), Path("master.xlsx"), snapshot, target)
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=write, args=(target,)) for target in targets]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert errors == []
    assert maximum_active_saves == 1


def test_verify_target_unchanged_rejects_a_changed_or_new_destination(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = Path("result.xlsx")
    old = FileSignature(size=3, mtime_ns=1, sha256="old")
    monkeypatch.setattr(Path, "exists", lambda _path: True)
    monkeypatch.setattr(Path, "is_file", lambda _path: True)
    monkeypatch.setattr("excel_splitter.excel_gateway._same_signature", lambda *_: False)

    with pytest.raises(OSError, match="변경"):
        _verify_target_unchanged(target, old)
    with pytest.raises(OSError, match="새로 생성"):
        _verify_target_unchanged(target, None)


def test_publish_new_target_uses_no_clobber_rename(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    temp = tmp_path / "temp.xlsx"
    target = tmp_path / "result.xlsx"
    temp.write_bytes(b"result")
    calls: list[tuple[Path, Path]] = []
    real_rename = os.rename

    def rename(source, destination):
        calls.append((Path(source), Path(destination)))
        return real_rename(source, destination)

    monkeypatch.setattr("excel_splitter.excel_gateway.os.rename", rename)
    _publish_temp(temp, target, None)

    assert calls == [(temp, target)]
    assert target.read_bytes() == b"result"


def test_publish_existing_target_claims_rechecks_and_removes_backup(
    tmp_path: Path,
) -> None:
    temp = tmp_path / "temp.xlsx"
    target = tmp_path / "result.xlsx"
    temp.write_bytes(b"new")
    target.write_bytes(b"old")
    prior = capture_signature(target)

    _publish_temp(temp, target, prior)

    assert target.read_bytes() == b"new"
    assert not tuple(tmp_path.glob(".esr-*.xlsx"))


def test_publish_changed_target_restores_claimed_file(tmp_path: Path) -> None:
    temp = tmp_path / "temp.xlsx"
    target = tmp_path / "result.xlsx"
    target.write_bytes(b"old")
    prior = capture_signature(target)
    target.write_bytes(b"changed")
    temp.write_bytes(b"new")

    with pytest.raises(OSError, match="변경"):
        _publish_temp(temp, target, prior)

    assert target.read_bytes() == b"changed"
    assert not tuple(tmp_path.glob(".esr-*.xlsx"))


def test_publish_backup_signature_read_failure_restores_target(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    temp = tmp_path / "temp.xlsx"
    target = tmp_path / "result.xlsx"
    target.write_bytes(b"old")
    prior = capture_signature(target)
    temp.write_bytes(b"new")
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._capture_signature",
        lambda _path: (_ for _ in ()).throw(OSError("read failed")),
    )

    with pytest.raises(OSError, match="검증하지 못했습니다"):
        _publish_temp(temp, target, prior)

    assert target.read_bytes() == b"old"
    assert not tuple(tmp_path.glob(".esr-*.xlsx"))


def test_publish_collision_preserves_recovery_without_overwriting_intruder(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    temp = tmp_path / "temp.xlsx"
    target = tmp_path / "result.xlsx"
    target.write_bytes(b"old")
    prior = capture_signature(target)
    target.write_bytes(b"changed")
    temp.write_bytes(b"new")
    real_rename = os.rename

    def rename(source, destination):
        source_path, destination_path = Path(source), Path(destination)
        if source_path.name.startswith(".esr-") and destination_path == target:
            target.write_bytes(b"intruder")
            raise FileExistsError("occupied")
        return real_rename(source, destination)

    monkeypatch.setattr("excel_splitter.excel_gateway.os.rename", rename)
    with pytest.raises(OSError, match="복구 파일을 보존") as captured:
        _publish_temp(temp, target, prior)

    recovery = tuple(tmp_path.glob(".esr-*.xlsx"))
    assert str(recovery[0]) in str(captured.value)
    assert recovery[0].read_bytes() == b"changed"
    assert target.read_bytes() == b"intruder"


def test_copy_to_master_removes_copy_when_signature_does_not_match(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = Path("source.xlsx")
    expected = FileSignature(size=6, mtime_ns=1, sha256="expected")
    observed = iter(
        (
            expected,
            expected,
            FileSignature(size=6, mtime_ns=1, sha256="different"),
        )
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._capture_signature", lambda _path: next(observed)
    )
    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", lambda *_args: None)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway.uuid.uuid4",
        lambda: SimpleNamespace(hex="token"),
    )
    removed: list[Path] = []
    monkeypatch.setattr(Path, "unlink", lambda path, **_kwargs: removed.append(path))

    with pytest.raises(SplitExecutionError, match="master"):
        _copy_to_master(source, expected, Path("output"))

    assert removed == [Path("output/.source.master.token.xlsx")]


def test_copy_to_master_is_created_in_the_target_parent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = Path("read-only/source.xlsx")
    output_parent = Path("writable-output")
    expected = FileSignature(size=6, mtime_ns=1, sha256="expected")
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._capture_signature", lambda _path: expected
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway.uuid.uuid4",
        lambda: SimpleNamespace(hex="token"),
    )
    copies: list[tuple[Path, Path]] = []
    monkeypatch.setattr(
        "excel_splitter.excel_gateway.shutil.copy2",
        lambda source_path, target_path: copies.append((source_path, target_path)),
    )

    master = _copy_to_master(source, expected, output_parent)

    assert master == Path("writable-output/.source.master.token.xlsx")
    assert copies == [(source, master)]


def test_write_groups_rejects_mixed_target_parents_before_copying(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source=Path("source.xlsx"),
        signature=FileSignature(1, 2, "abc"),
        sheet_name="Sheet1",
        table_name="Table1",
        column_name="Team",
        row_count=1,
        groups=(GroupSummary(key, "A", 1, (1,)),),
    )
    targets = (
        OutputTarget(key, "A", Path("one/a.xlsx"), None),
        OutputTarget(key, "A", Path("two/b.xlsx"), None),
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._copy_to_master",
        lambda *_args: pytest.fail("master copy must not start"),
    )

    with pytest.raises(SplitExecutionError, match="같은 출력 폴더"):
        ExcelComGateway().write_groups(snapshot, targets, lambda *_args: None)


def test_group_cleanup_does_not_mask_a_com_trust_error_with_oserror(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source=Path("source.xlsx"),
        signature=FileSignature(1, 2, "abc"),
        sheet_name="Sheet1",
        table_name="Table1",
        column_name="Team",
        row_count=1,
        groups=(GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", Path("result.xlsx"), None)

    class UncloseableWorkbook:
        def Close(self, **_kwargs) -> None:
            raise RuntimeError("close failed")

    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", lambda *_: None)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._open_workbook",
        lambda *_args, **_kwargs: UncloseableWorkbook(),
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda *_args: (_ for _ in ()).throw(SplitExecutionError("identity failed")),
    )
    monkeypatch.setattr(
        Path,
        "unlink",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("temp locked")),
    )

    with pytest.raises(SplitExecutionError, match="identity failed"):
        _write_one_group(SimpleNamespace(), Path("master.xlsx"), snapshot, target)


def test_master_cleanup_preserves_trust_error_and_aborts_remaining_groups(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source=tmp_path / "source.xlsx",
        signature=FileSignature(1, 2, "abc"),
        sheet_name="Sheet1",
        table_name="Table1",
        column_name="Team",
        row_count=1,
        groups=(GroupSummary(key, "A", 1, (1,)),),
    )
    targets = (
        OutputTarget(key, "first", tmp_path / "first.xlsx", None),
        OutputTarget(key, "second", tmp_path / "second.xlsx", None),
    )
    class Session:
        def start(self): pass
        def open_source(self, _source): return SimpleNamespace(sheets=("Sheet1",))
        def save_plain_master(self, run_dir, _snapshot):
            master = run_dir / "master.xlsx"
            master.write_bytes(b"master")
            return master
        def close_source(self): pass
        def shutdown(self): pass

    monkeypatch.setattr(
        "excel_splitter.excel_gateway.write_targets",
        lambda *_args: (_ for _ in ()).throw(SplitExecutionError("session untrusted")),
    )
    cleanup_attempts: list[Path] = []
    monkeypatch.setattr(
        "excel_splitter.excel_gateway.shutil.rmtree",
        lambda path, **_kwargs: (
            cleanup_attempts.append(path),
            (_ for _ in ()).throw(OSError("run dir locked")),
        )[1],
    )
    monkeypatch.setattr("excel_splitter.excel_gateway.time.sleep", lambda _delay: None)

    with pytest.raises(SplitExecutionError, match="session untrusted.*평문") as captured:
        ExcelComGateway(Session()).write_groups(snapshot, targets, lambda *_args: None)

    assert len(cleanup_attempts) == 3
    assert str(cleanup_attempts[0]) in str(captured.value)


def test_gateway_reuses_source_session_and_runs_plain_master_before_parallel_write(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source=source,
        signature=FileSignature(6, 1, "abc"),
        sheet_name="Sheet1",
        table_name="Table1",
        column_name="Team",
        row_count=1,
        groups=(GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", tmp_path / "A.xlsx", None)
    events: list[object] = []

    class Session:
        def start(self):
            events.append("start")

        def open_source(self, path):
            events.append(("open", path))
            return SimpleNamespace(sheets=("Sheet1",))

        def inspect_table(self, sheet_name):
            events.append(("inspect", sheet_name))
            return TableInfo(sheet_name, "Table1", ("Team",), 1)

        def build_snapshot(self, sheet_name, column_name):
            events.append(("snapshot", sheet_name, column_name))
            return snapshot

        def save_plain_master(self, run_dir, actual_snapshot):
            events.append(("master", actual_snapshot))
            master = run_dir / "master.xlsx"
            master.write_bytes(b"master")
            return master

        def close_source(self):
            events.append("close")

        def shutdown(self):
            events.append("shutdown")

    def parallel(master, actual_snapshot, targets, write_one, session_factory, progress):
        assert session_factory is excel_gateway._output_excel_session
        events.append(("parallel", master.exists(), actual_snapshot, targets))
        return SplitResult((targets[0].path,), ())

    monkeypatch.setattr("excel_splitter.excel_gateway.write_targets", parallel)
    gateway = ExcelComGateway(Session())

    assert gateway.list_worksheets(source) == ("Sheet1",)
    assert gateway.inspect_table(source, "Sheet1").columns == ("Team",)
    assert gateway.build_snapshot(source, "Sheet1", "Team") is snapshot
    result = gateway.write_groups(snapshot, (target,), lambda *_: None)
    gateway.shutdown()

    assert result.succeeded == (target.path,)
    assert events[:4] == [
        "start",
        ("open", source),
        ("inspect", "Sheet1"),
        ("snapshot", "Sheet1", "Team"),
    ]
    master_index = next(i for i, event in enumerate(events) if isinstance(event, tuple) and event[0] == "master")
    close_index = events.index("close")
    parallel_index = next(i for i, event in enumerate(events) if isinstance(event, tuple) and event[0] == "parallel")
    assert master_index < close_index < parallel_index
    assert events[-1] == "shutdown"
    assert not any(path.name.startswith(".es-") for path in tmp_path.iterdir())


def test_gateway_reopens_when_source_changes_and_clears_cache_after_write(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    first.write_bytes(b"first")
    second.write_bytes(b"second")
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        first, FileSignature(5, 1, "hash"), "S", "T", "C", 1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    events: list[object] = []

    class Session:
        def start(self): pass
        def open_source(self, path):
            events.append(("open", path))
            return SimpleNamespace(sheets=("S",))
        def inspect_table(self, name): return TableInfo(name, "T", ("C",), 1)
        def build_snapshot(self, *_args): return snapshot
        def save_plain_master(self, run_dir, _snapshot):
            path = run_dir / "master.xlsx"
            path.write_bytes(b"master")
            return path
        def close_source(self): events.append("close")
        def shutdown(self): pass

    monkeypatch.setattr(
        "excel_splitter.excel_gateway.write_targets",
        lambda _master, _snapshot, targets, *_args: SplitResult(
            tuple(t.path for t in targets), ()
        ),
    )
    gateway = ExcelComGateway(Session())
    gateway.list_worksheets(first)
    gateway.inspect_table(first, "S")
    gateway.build_snapshot(first, "S", "C")
    gateway.inspect_table(second, "S")
    target = OutputTarget(key, "A", tmp_path / "A.xlsx", None)
    gateway.write_groups(snapshot, (target,), lambda *_: None)
    gateway.inspect_table(first, "S")

    assert [event for event in events if isinstance(event, tuple)] == [
        ("open", first), ("open", second), ("open", first), ("open", first)
    ]


def test_same_preview_can_execute_twice_by_reopening_source(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source, FileSignature(6, 1, "hash"), "S", "T", "C", 1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", tmp_path / "A.xlsx", None)
    opens: list[Path] = []

    class Session:
        def start(self): pass
        def open_source(self, path):
            opens.append(path); return SimpleNamespace(sheets=("S",))
        def save_plain_master(self, run_dir, _snapshot):
            master = run_dir / "m.xlsx"; master.write_bytes(b"master"); return master
        def close_source(self): pass
        def shutdown(self): pass

    monkeypatch.setattr(
        "excel_splitter.excel_gateway.write_targets",
        lambda _master, _snapshot, targets, *_args: SplitResult(
            tuple(item.path for item in targets), ()
        ),
    )
    gateway = ExcelComGateway(Session())

    gateway.write_groups(snapshot, (target,), lambda *_: None)
    gateway.write_groups(snapshot, (target,), lambda *_: None)

    assert opens == [source, source]


def test_write_groups_rejects_excel_temp_path_over_218_characters(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        source, FileSignature(6, 1, "hash"), "S", "T", "C", 1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", tmp_path / "A.xlsx", None)
    fake_run_dir = tmp_path / ("x" * 210)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway.tempfile.mkdtemp", lambda **_kwargs: str(fake_run_dir)
    )
    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.rmtree", lambda *_args: None)

    class Session:
        def start(self): pass
        def shutdown(self): pass

    with pytest.raises(SplitExecutionError, match="218"):
        ExcelComGateway(Session()).write_groups(snapshot, (target,), lambda *_: None)


def test_gateway_invalidates_cached_source_after_inspection_failure() -> None:
    source = Path("source.xlsx")
    opens: list[Path] = []

    class Session:
        def start(self): pass
        def open_source(self, path):
            opens.append(path)
            return SimpleNamespace(sheets=("S",))
        def inspect_table(self, _name): raise WorkbookValidationError("changed")
        def shutdown(self): pass

    gateway = ExcelComGateway(Session())
    gateway.list_worksheets(source)
    with pytest.raises(WorkbookValidationError, match="changed"):
        gateway.inspect_table(source, "S")
    with pytest.raises(WorkbookValidationError, match="changed"):
        gateway.inspect_table(source, "S")

    assert opens == [source, source]


def test_gateway_closes_source_after_non_validation_inspection_failure() -> None:
    events: list[object] = []

    class Session:
        def start(self): pass
        def open_source(self, path):
            events.append(("open", path)); return SimpleNamespace(sheets=("S",))
        def inspect_table(self, _name): raise RuntimeError("COM trust lost")
        def close_source(self): events.append("close")
        def shutdown(self): pass

    source = Path("source.xlsx")
    gateway = ExcelComGateway(Session())
    gateway.list_worksheets(source)
    with pytest.raises(RuntimeError, match="trust lost"):
        gateway.inspect_table(source, "S")

    assert events == [("open", source), "close"]


def test_failed_open_cannot_leave_a_stale_gateway_cache() -> None:
    first, second = Path("first.xlsx"), Path("second.xlsx")
    opens: list[Path] = []

    class Session:
        def start(self): pass
        def open_source(self, path):
            opens.append(path)
            if path == second:
                raise RuntimeError("open failed")
            return SimpleNamespace(sheets=("S",))
        def inspect_table(self, name): return TableInfo(name, "T", ("C",), 1)
        def shutdown(self): pass

    gateway = ExcelComGateway(Session())
    gateway.list_worksheets(first)
    with pytest.raises(RuntimeError, match="open failed"):
        gateway.inspect_table(second, "S")
    gateway.inspect_table(first, "S")

    assert opens == [first, second, first]


def test_group_temp_copy_is_kept_in_master_run_directory(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"), FileSignature(1, 2, "abc"), "S", "T", "C", 1,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    target = OutputTarget(key, "A", Path("output/result.xlsx"), None)
    copies: list[tuple[Path, Path]] = []
    monkeypatch.setattr(
        "excel_splitter.excel_gateway.shutil.copy2",
        lambda source, destination: copies.append((source, destination)),
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._open_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(SplitExecutionError("stop")),
    )
    monkeypatch.setattr(Path, "unlink", lambda *_args, **_kwargs: None)

    with pytest.raises(SplitExecutionError, match="stop"):
        _write_one_group(SimpleNamespace(), Path("run/master.xlsx"), snapshot, target)

    assert copies[0][1].parent == Path("run")
    assert copies[0][1].name.startswith("g-")
    assert len(copies[0][1].name) <= 19
