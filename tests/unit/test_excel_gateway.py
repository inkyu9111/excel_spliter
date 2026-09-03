from __future__ import annotations

import sys
import os
import shutil
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
    _com_stage,
    _contiguous_descending_blocks,
    _copy_to_master,
    _delete_row_blocks,
    _delete_rows,
    _excel_error_code,
    _excel_session,
    _open_workbook,
    _publish_temp,
    _remove_other_sheets,
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


@pytest.mark.parametrize(
    ("indexes", "expected"),
    (
        ((), ()),
        ((4, 2, 3, 2, 0, -1), ((2, 4),)),
        ((1, 2, 5, 7, 6, 10), ((10, 10), (5, 7), (1, 2))),
    ),
)
def test_contiguous_descending_blocks_normalize_and_coalesce(
    indexes: tuple[int, ...], expected: tuple[tuple[int, int], ...]
) -> None:
    assert _contiguous_descending_blocks(indexes) == expected


def test_delete_row_blocks_uses_table_width_ranges_in_descending_order() -> None:
    deleted: list[tuple[int, int, int]] = []

    class Worksheet:
        def Range(self, first, last):
            return SimpleNamespace(
                Delete=lambda **options: deleted.append(
                    (first.index, last.index, options["Shift"])
                )
            )

    worksheet = Worksheet()

    class Rows:
        def __call__(self, index: int):
            return SimpleNamespace(
                Range=SimpleNamespace(index=index, Worksheet=worksheet)
            )

    _delete_row_blocks(Rows(), (1, 2, 5, 7, 6, 10))

    assert deleted == [(10, 10, -4162), (5, 7, -4162), (1, 2, -4162)]


def test_delete_row_blocks_signals_excel_1004_as_compatible() -> None:
    error = RuntimeError(-2147352567, "Delete method failed")
    error.hresult = -2147352567  # type: ignore[attr-defined]
    error.excepinfo = (  # type: ignore[attr-defined]
        0,
        "Microsoft Excel",
        "Delete method failed",
        None,
        0,
        -2146827284,
    )

    class Worksheet:
        @staticmethod
        def Range(_first, _last):
            return SimpleNamespace(
                Delete=lambda **_options: (_ for _ in ()).throw(error)
            )

    row_range = SimpleNamespace(Worksheet=Worksheet())
    rows = lambda _index: SimpleNamespace(Range=row_range)

    with pytest.raises(excel_gateway._BlockDeleteCompatibilityError) as caught:
        _delete_row_blocks(rows, (1,))

    assert caught.value.__cause__ is error


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


def test_com_stage_includes_hresult_and_excepinfo_without_losing_cause() -> None:
    error = FakeComError(-2147352567, "예외가 발생했습니다.", -2147352561)

    with pytest.raises(SplitExecutionError) as captured:
        with _com_stage("조건부 서식 추가: Data R7C4"):
            raise error

    message = str(captured.value)
    assert "조건부 서식 추가: Data R7C4" in message
    assert "0x80020009" in message and "0x8002000F" in message
    assert "필수 매개 변수" in message
    assert captured.value.__cause__ is error


def test_output_excel_session_does_not_access_calculation_before_workbook_open(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    created: list[object] = []

    class Excel:
        @property
        def Calculation(self):
            raise AssertionError("Calculation accessed before workbook open")

        @Calculation.setter
        def Calculation(self, _value) -> None:
            raise AssertionError("Calculation changed before workbook open")

        @property
        def CalculateBeforeSave(self):
            raise AssertionError("CalculateBeforeSave accessed before workbook open")

        @CalculateBeforeSave.setter
        def CalculateBeforeSave(self, _value) -> None:
            raise AssertionError("CalculateBeforeSave changed before workbook open")

        def Quit(self) -> None:
            pass

    def dispatch(_name: str) -> Excel:
        excel = Excel()
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
        assert output_excel is created[0]


@pytest.mark.parametrize(
    ("failed_stage", "expected_message"),
    (
        ("open", "파일 열기"),
        ("delete", "행 삭제"),
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
            "excel_splitter.excel_gateway._delete_row_blocks",
            lambda *_: (_ for _ in ()).throw(RuntimeError("COM delete failed")),
        )
    with pytest.raises(SplitExecutionError, match=expected_message):
        _write_one_group(
            SimpleNamespace(Calculate=lambda: None),
            Path("master.xlsx"),
            snapshot,
            target,
        )


def _stub_successful_group_write(
    monkeypatch: pytest.MonkeyPatch,
    open_workbook,
    delete_rows=lambda *_args: None,
    sheet=None,
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
    if sheet is None:
        sheet = SimpleNamespace(FilterMode=False)
    table = SimpleNamespace(Name="Table1", ListRows=SimpleNamespace(Count=1))
    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", lambda *_: None)
    monkeypatch.setattr("excel_splitter.excel_gateway._open_workbook", open_workbook)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda *_args: (sheet, table),
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr("excel_splitter.excel_gateway._delete_row_blocks", delete_rows)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._publish_temp", lambda *_: None)
    return snapshot, target


def test_write_one_group_retries_compatible_block_failure_from_fresh_master_copy(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    master = tmp_path / "m.xlsx"
    master.write_bytes(b"clean master")
    target = OutputTarget(key, "A", tmp_path / "result.xlsx", None)
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Sheet1",
        "Table1",
        "Team",
        3,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    events: list[object] = []

    class Rows:
        def __init__(self, *, supports_ranges: bool) -> None:
            self.Count = 3
            self.supports_ranges = supports_ranges

        def __call__(self, index: int):
            if not self.supports_ranges:
                return SimpleNamespace()

            owner = self

            class Row:
                def Delete(self) -> None:
                    events.append(("row.delete", index))
                    owner.Count -= 1

            return Row()

    class Workbook:
        Sheets = SimpleNamespace(Count=1)

        def __init__(self, attempt: int) -> None:
            self.attempt = attempt
            self.sheet = SimpleNamespace(FilterMode=False)
            self.table = SimpleNamespace(
                Name="Table1", ListRows=Rows(supports_ranges=attempt == 2)
            )

        def Save(self) -> None:
            events.append(("save", self.attempt))

        def Close(self, **options) -> None:
            events.append(("close", self.attempt, options["SaveChanges"]))

    opened: list[Workbook] = []

    def open_workbook(*_args, **_kwargs):
        workbook = Workbook(len(opened) + 1)
        opened.append(workbook)
        events.append(("open", workbook.attempt))
        return workbook

    copies: list[tuple[Path, Path]] = []
    destination_existed_before_copy: list[bool] = []
    real_copy2 = shutil.copy2

    def copy2(source: Path, destination: Path):
        copies.append((Path(source), Path(destination)))
        destination_existed_before_copy.append(Path(destination).exists())
        return real_copy2(source, destination)

    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", copy2)
    monkeypatch.setattr("excel_splitter.excel_gateway._open_workbook", open_workbook)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda workbook, _sheet_name: (workbook.sheet, workbook.table),
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._publish_temp",
        lambda temp, output, _signature: events.append(("publish", temp, output)),
    )
    excel = SimpleNamespace(
        Calculation=-4105,
        CalculateBeforeSave=False,
        Calculate=lambda: events.append("calculate"),
    )

    assert _write_one_group(excel, master, snapshot, target) == target.path

    assert len(copies) == 2
    assert copies[0] == copies[1]
    assert destination_existed_before_copy == [False, False]
    assert [("open", 1), ("close", 1, False), ("open", 2)] == [
        event for event in events if event[0] in {"open", "close"}
    ][:3]
    assert [event for event in events if event[0] == "row.delete"] == [
        ("row.delete", 3),
        ("row.delete", 2),
    ]
    assert len([event for event in events if event[0] == "publish"]) == 1


def test_write_one_group_propagates_fatal_rowwise_retry_and_cleans_both_attempts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    key = CanonicalKey("text", "A")
    master = tmp_path / "m.xlsx"
    master.write_bytes(b"clean master")
    target = OutputTarget(key, "A", tmp_path / "result.xlsx", None)
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Sheet1",
        "Table1",
        "Team",
        3,
        (GroupSummary(key, "A", 1, (1,)),),
    )
    events: list[object] = []

    class Rows:
        Count = 3

        def __init__(self, attempt: int) -> None:
            self.attempt = attempt

        def __call__(self, index: int):
            if self.attempt == 1:
                return SimpleNamespace()
            return SimpleNamespace(
                Delete=lambda: (_ for _ in ()).throw(
                    RuntimeError(f"row-wise failure at {index}")
                )
            )

    class Workbook:
        Sheets = SimpleNamespace(Count=1)

        def __init__(self, attempt: int) -> None:
            self.attempt = attempt
            self.sheet = SimpleNamespace(FilterMode=False)
            self.table = SimpleNamespace(
                Name="Table1", ListRows=Rows(attempt)
            )

        def Save(self) -> None:
            events.append(("save", self.attempt))

        def Close(self, **options) -> None:
            events.append(("close", self.attempt, options["SaveChanges"]))

    opened: list[Workbook] = []

    def open_workbook(*_args, **_kwargs):
        workbook = Workbook(len(opened) + 1)
        opened.append(workbook)
        events.append(("open", workbook.attempt))
        return workbook

    copies: list[tuple[Path, Path]] = []
    real_copy2 = shutil.copy2

    def copy2(source: Path, destination: Path):
        copies.append((Path(source), Path(destination)))
        return real_copy2(source, destination)

    monkeypatch.setattr("excel_splitter.excel_gateway.shutil.copy2", copy2)
    monkeypatch.setattr("excel_splitter.excel_gateway._open_workbook", open_workbook)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._validated_table",
        lambda workbook, _sheet_name: (workbook.sheet, workbook.table),
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._publish_temp",
        lambda *_args: events.append("publish"),
    )

    with pytest.raises(SplitExecutionError, match="행 삭제") as captured:
        _write_one_group(
            SimpleNamespace(
                Calculation=-4105,
                CalculateBeforeSave=False,
                Calculate=lambda: events.append("calculate"),
            ),
            master,
            snapshot,
            target,
        )

    assert "row-wise failure" in str(captured.value.__cause__)
    assert len(copies) == 2
    assert len(opened) == 2
    assert [event for event in events if isinstance(event, tuple) and event[0] == "close"] == [
        ("close", 1, False),
        ("close", 2, False),
    ]
    assert "publish" not in events
    assert not target.path.exists()
    assert tuple(tmp_path.glob("g-*.xlsx")) == ()


def test_write_one_group_does_not_retry_unknown_block_delete_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    master = tmp_path / "m.xlsx"
    master.write_bytes(b"clean master")
    opens = 0

    def open_workbook(*_args, **_kwargs):
        nonlocal opens
        opens += 1
        return SimpleNamespace(
            Sheets=SimpleNamespace(Count=1),
            Save=lambda: None,
            Close=lambda **_options: None,
        )

    snapshot, target = _stub_successful_group_write(monkeypatch, open_workbook)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._delete_row_blocks",
        lambda *_args: (_ for _ in ()).throw(RuntimeError("transport lost")),
    )

    with pytest.raises(SplitExecutionError, match="행 삭제"):
        _write_one_group(
            SimpleNamespace(
                Calculation=-4105,
                CalculateBeforeSave=False,
                Calculate=lambda: None,
            ),
            master,
            snapshot,
            target,
        )

    assert opens == 1


def test_write_one_group_never_restores_comment_shape_by_name(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    shape = SimpleNamespace(
        Name="Comment 4",
        Left=10.0,
        Top=20.0,
        Width=30.0,
        Height=40.0,
        Placement=1,
    )
    lookups: list[object] = []

    class Shapes:
        Count = 1

        def Item(self, key: object):
            lookups.append(key)
            return shape

    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: None,
        Close=lambda **_kwargs: None,
    )
    sheet = SimpleNamespace(FilterMode=False, Shapes=Shapes())
    snapshot, target = _stub_successful_group_write(
        monkeypatch,
        lambda *_args, **_kwargs: workbook,
        sheet=sheet,
    )

    assert (
        _write_one_group(
            SimpleNamespace(Calculate=lambda: None),
            Path("master.xlsx"),
            snapshot,
            target,
        )
        == target.path
    )
    assert lookups == []


def test_write_one_group_reapplies_manual_calculation_after_workbook_open(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[tuple[str, int, bool]] = []

    class Excel:
        Calculation = -4135
        CalculateBeforeSave = True

        def Calculate(self) -> None:
            events.append(
                ("calculate", self.Calculation, self.CalculateBeforeSave)
            )

    excel = Excel()

    class Workbook:
        Sheets = SimpleNamespace(Count=1)

        def Save(self) -> None:
            events.append(("save", excel.Calculation, excel.CalculateBeforeSave))

        def Close(self, **_kwargs) -> None:
            pass

    workbook = Workbook()
    observed_before_delete: list[tuple[int, bool]] = []

    def open_workbook(*_args, **_kwargs):
        excel.Calculation = -4105
        excel.CalculateBeforeSave = False
        return workbook

    def delete_rows(*_args):
        observed_before_delete.append(
            (excel.Calculation, excel.CalculateBeforeSave)
        )

    snapshot, target = _stub_successful_group_write(
        monkeypatch, open_workbook, delete_rows
    )

    _write_one_group(excel, Path("master.xlsx"), snapshot, target)

    assert observed_before_delete == [(-4135, True)]
    assert events == [
        ("calculate", -4135, True),
        ("save", -4105, False),
    ]
    assert (excel.Calculation, excel.CalculateBeforeSave) == (-4105, False)


def test_write_one_group_falls_back_when_calculation_settings_do_not_read_back(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class Excel:
        def __init__(self, events: list[str]) -> None:
            self._calculation = -4105
            self._calculate_before_save = False
            self.force_readback_mismatch = False
            self._events = events

        def Calculate(self) -> None:
            self._events.append("application.calculate")

        @property
        def Calculation(self):
            return self._calculation

        @Calculation.setter
        def Calculation(self, value) -> None:
            self._calculation = value

        @property
        def CalculateBeforeSave(self):
            if self.force_readback_mismatch and self._calculation == -4135:
                return False
            return self._calculate_before_save

        @CalculateBeforeSave.setter
        def CalculateBeforeSave(self, value) -> None:
            self._calculate_before_save = value

    events: list[str] = []
    excel = Excel(events)
    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: None,
        Close=lambda **_kwargs: None,
    )

    def open_workbook(*_args, **_kwargs):
        excel.Calculation = -4105
        excel.CalculateBeforeSave = False
        excel.force_readback_mismatch = True
        return workbook

    snapshot, target = _stub_successful_group_write(monkeypatch, open_workbook)

    assert _write_one_group(excel, Path("master.xlsx"), snapshot, target) == target.path
    assert events == ["application.calculate"]
    assert excel.Calculation == -4105
    assert excel.CalculateBeforeSave is False


def test_manual_calculation_allows_parallel_mutation_but_serializes_saves(
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
    active_mutations = 0
    maximum_active_mutations = 0
    mutation_barrier = threading.Barrier(2)

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

    sheet = SimpleNamespace(FilterMode=False)
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

    def delete_rows(*_args):
        nonlocal active_mutations, maximum_active_mutations
        with state_lock:
            active_mutations += 1
            maximum_active_mutations = max(maximum_active_mutations, active_mutations)
        mutation_barrier.wait(timeout=1)
        with state_lock:
            active_mutations -= 1

    monkeypatch.setattr("excel_splitter.excel_gateway._delete_row_blocks", delete_rows)
    monkeypatch.setattr("excel_splitter.excel_gateway._column_index", lambda *_: 1)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._publish_temp", lambda *_: None)
    errors: list[Exception] = []

    def write(target: OutputTarget) -> None:
        try:
            _write_one_group(
                SimpleNamespace(Calculate=lambda: None),
                Path("master.xlsx"),
                snapshot,
                target,
            )
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=write, args=(target,)) for target in targets]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert errors == []
    assert maximum_active_mutations == 2
    assert maximum_active_saves == 1


class FakeComError(Exception):
    def __init__(self, hresult: int, message: str, scode: int | None = None) -> None:
        excepinfo = (0, "Microsoft Excel", message, None, 0, scode)
        super().__init__(hresult, message, excepinfo, None)
        self.hresult = hresult
        self.excepinfo = excepinfo


def test_manual_calculation_restore_failure_aborts_before_save_and_publish(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    class Excel:
        def __init__(self) -> None:
            self._calculation = -4105
            self.CalculateBeforeSave = False

        @property
        def Calculation(self):
            return self._calculation

        @Calculation.setter
        def Calculation(self, value) -> None:
            if value == -4105 and self._calculation == -4135:
                raise FakeComError(
                    -2147352567, "Unable to restore Calculation", -2146827284
                )
            self._calculation = value

        def Calculate(self) -> None:
            events.append("calculate")

    excel = Excel()
    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: events.append("save"),
        Close=lambda **_kwargs: None,
    )
    snapshot, target = _stub_successful_group_write(
        monkeypatch, lambda *_args, **_kwargs: workbook
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._publish_temp",
        lambda *_args: events.append("publish"),
    )

    with pytest.raises(SplitExecutionError, match="저장"):
        _write_one_group(excel, Path("master.xlsx"), snapshot, target)

    assert events == ["calculate"]


def test_partial_calculation_restore_failure_is_fatal_before_mutation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    class Excel:
        def __init__(self) -> None:
            self._calculation = -4105
            self._calculate_before_save = False

        @property
        def Calculation(self):
            return self._calculation

        @Calculation.setter
        def Calculation(self, value) -> None:
            if value == -4105 and self._calculation == -4135:
                raise FakeComError(
                    -2147352567, "Unable to restore Calculation", -2146827284
                )
            self._calculation = value

        @property
        def CalculateBeforeSave(self):
            if self._calculation == -4135:
                return False
            return self._calculate_before_save

        @CalculateBeforeSave.setter
        def CalculateBeforeSave(self, value) -> None:
            self._calculate_before_save = value

        def Calculate(self) -> None:
            events.append("calculate")

    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: events.append("save"),
        Close=lambda **_kwargs: None,
    )
    snapshot, target = _stub_successful_group_write(
        monkeypatch, lambda *_args, **_kwargs: workbook,
        delete_rows=lambda *_args: events.append("mutate"),
    )

    with pytest.raises(SplitExecutionError, match="파일 열기"):
        _write_one_group(Excel(), Path("master.xlsx"), snapshot, target)

    assert events == []


def test_calculation_1004_fallback_serializes_mutation_and_save(
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
    active_operations = 0
    maximum_active_operations = 0

    class Excel:
        def Calculate(self) -> None:
            pass

        @property
        def Calculation(self):
            return -4105

        @Calculation.setter
        def Calculation(self, _value) -> None:
            raise FakeComError(
                -2147352567,
                "Unable to set the Calculation property of the Application class",
                -2146827284,
            )

    class Workbook:
        Sheets = SimpleNamespace(Count=1)

        def Save(self) -> None:
            nonlocal active_operations
            time.sleep(0.03)
            with state_lock:
                active_operations -= 1

        def Close(self, **_kwargs) -> None:
            pass

    sheet = SimpleNamespace(FilterMode=False)
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

    def delete_rows(*_args):
        nonlocal active_operations, maximum_active_operations
        with state_lock:
            active_operations += 1
            maximum_active_operations = max(
                maximum_active_operations, active_operations
            )

    monkeypatch.setattr("excel_splitter.excel_gateway._delete_row_blocks", delete_rows)
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._remove_other_sheets", lambda *_: None
    )
    monkeypatch.setattr("excel_splitter.excel_gateway._publish_temp", lambda *_: None)
    errors: list[Exception] = []

    def write(target: OutputTarget) -> None:
        try:
            _write_one_group(Excel(), Path("master.xlsx"), snapshot, target)
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=write, args=(target,)) for target in targets]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert errors == []
    assert maximum_active_operations == 1


def test_noncompatibility_com_error_while_setting_calculation_is_fatal(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class Excel:
        @property
        def Calculation(self):
            return -4105

        @Calculation.setter
        def Calculation(self, _value) -> None:
            raise FakeComError(-2147417848, "The object invoked has disconnected")

    workbook = SimpleNamespace(
        Sheets=SimpleNamespace(Count=1),
        Save=lambda: None,
        Close=lambda **_kwargs: None,
    )
    snapshot, target = _stub_successful_group_write(
        monkeypatch, lambda *_args, **_kwargs: workbook
    )

    with pytest.raises(SplitExecutionError, match="파일 열기") as captured:
        _write_one_group(Excel(), Path("master.xlsx"), snapshot, target)

    assert isinstance(captured.value.__cause__, FakeComError)
    assert captured.value.__cause__.hresult == -2147417848


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
