from __future__ import annotations

import sys
import gc
from contextlib import contextmanager, nullcontext
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_splitter.errors import (
    SplitExecutionError,
    WorkbookValidationError,
)
from excel_splitter.models import (
    CanonicalKey,
    FileSignature,
    GroupSummary,
    OutputTarget,
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
    _remove_other_sheets,
    _single_table,
    _verify_target_unchanged,
    _write_one_group,
)


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


def test_list_worksheets_releases_workbook_before_session_exit(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    class Workbook:
        Worksheets = SimpleNamespace(Count=0)

        def Close(self, **_kwargs) -> None:
            events.append("close")

        def __del__(self) -> None:
            events.append("release")

    excel = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=lambda *_args, **_kwargs: Workbook())
    )

    @contextmanager
    def session():
        yield excel
        gc.collect()
        events.append("session_exit")

    monkeypatch.setattr("excel_splitter.excel_gateway._excel_session", session)

    assert ExcelComGateway().list_worksheets(Path("source.xlsx")) == ()
    assert events.index("release") < events.index("session_exit")


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
        OutputTarget(key, "first", Path("output/first.xlsx"), None),
        OutputTarget(key, "second", Path("output/second.xlsx"), None),
    )
    calls: list[str] = []
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._copy_to_master",
        lambda *_args: Path("output/master.xlsx"),
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._excel_session",
        lambda: nullcontext(SimpleNamespace()),
    )
    monkeypatch.setattr(
        "excel_splitter.excel_gateway._write_one_group",
        lambda _excel, _master, _snapshot, target: (
            calls.append(target.label),
            (_ for _ in ()).throw(SplitExecutionError("session untrusted")),
        )[1],
    )
    monkeypatch.setattr(
        Path,
        "unlink",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("master locked")),
    )

    with pytest.raises(SplitExecutionError, match="session untrusted"):
        ExcelComGateway().write_groups(snapshot, targets, lambda *_args: None)

    assert calls == ["first"]
