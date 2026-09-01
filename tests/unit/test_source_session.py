from __future__ import annotations

import threading
import zipfile
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_splitter.errors import SplitExecutionError, WorkbookValidationError
from excel_splitter.models import CanonicalKey, GroupSummary
from excel_splitter.source_session import SourceHandleInfo, SourceSession


class Collection:
    def __init__(self, items: list[object]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, index: int):
        return self._items[index - 1]


class Workbook:
    def __init__(
        self,
        events: list[tuple[str, int]],
        *,
        sheet_name: str = "Data",
        table_name: str = "Orders",
        columns: tuple[str, ...] = ("Team", "Amount"),
        row_count: int = 2,
    ) -> None:
        self.events = events
        data_body_range = SimpleNamespace(
            Value2=tuple(("A",) for _ in range(row_count)),
            Cells=SimpleNamespace(
                Item=lambda _row, _column: SimpleNamespace(Text="A")
            ),
        )
        list_columns = [
            SimpleNamespace(Name=name, DataBodyRange=data_body_range)
            for name in columns
        ]
        self.table = SimpleNamespace(
            Name=table_name,
            SourceType=1,
            ListRows=SimpleNamespace(Count=row_count),
            ListColumns=Collection(list_columns),
            Range=SimpleNamespace(
                Row=1,
                Rows=SimpleNamespace(Count=row_count + 1),
                Column=1,
                Columns=SimpleNamespace(Count=len(columns)),
            ),
        )
        self.sheet = SimpleNamespace(
            Name=sheet_name,
            Visible=-1,
            ProtectContents=False,
            ListObjects=Collection([self.table]),
            UsedRange=SimpleNamespace(
                Row=1, Rows=SimpleNamespace(Count=row_count + 1)
            ),
            Comments=SimpleNamespace(Count=0),
            CommentsThreaded=SimpleNamespace(Count=0),
            Shapes=SimpleNamespace(Count=0),
        )
        self.Worksheets = Collection([self.sheet])
        self.ProtectStructure = False

    def Close(self, **options: object) -> None:
        assert options == {"SaveChanges": False}
        self.events.append(("close", threading.get_ident()))


class Excel:
    def __init__(
        self,
        events: list[tuple[str, int]],
        opener,
    ) -> None:
        self.events = events
        self.open_calls: list[tuple[str, dict[str, object], int]] = []

        def open_workbook(source: str, **options: object):
            self.open_calls.append((source, options, threading.get_ident()))
            return opener(Path(source))

        self.Workbooks = SimpleNamespace(Open=open_workbook)

    def Quit(self) -> None:
        self.events.append(("quit", threading.get_ident()))


def session_factory_for(excel: Excel, events: list[tuple[str, int]]):
    @contextmanager
    def factory():
        events.append(("initialize", threading.get_ident()))
        try:
            yield excel
        finally:
            excel.Quit()
            events.append(("uninitialize", threading.get_ident()))

    return factory


def write_minimal_xlsx(path: Path) -> None:
    entries = (
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
    )
    with zipfile.ZipFile(path, "w") as package:
        for name in entries:
            package.writestr(name, "<xml />")


def test_open_inspect_and_snapshot_reuse_one_thread_owned_workbook(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    excel = Excel(events, lambda _path: workbook)
    group = GroupSummary(CanonicalKey("text", "A"), "A", 2, (1, 2))
    sample_threads: list[int] = []

    def group_bulk_samples(_range: object, row_count: int):
        assert row_count == 2
        sample_threads.append(threading.get_ident())
        return (group,)

    monkeypatch.setattr(
        "excel_splitter.source_session._group_bulk_samples", group_bulk_samples
    )
    caller_thread = threading.get_ident()
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    handle = session.open_source(source)
    info = session.inspect_table("Data")
    snapshot = session.build_snapshot("Data", "Team")
    session.shutdown()

    assert isinstance(handle, SourceHandleInfo)
    assert handle.source == source
    assert handle.sheets == ("Data",)
    assert info.table_name == "Orders"
    assert snapshot.groups == (group,)
    assert snapshot.has_removable_artifacts is False
    assert len(excel.open_calls) == 1
    worker_threads = {thread_id for _, _, thread_id in excel.open_calls}
    worker_threads.update(sample_threads)
    worker_threads.update(thread_id for _, thread_id in events)
    assert len(worker_threads) == 1
    assert caller_thread not in worker_threads


def test_inspect_and_snapshot_use_quick_stat_without_rehashing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    excel = Excel(events, lambda _path: workbook)
    from excel_splitter.file_signature import capture_signature as real_capture

    captures: list[Path] = []

    def capture(path: Path):
        captures.append(path)
        return real_capture(path)

    monkeypatch.setattr("excel_splitter.source_session.capture_signature", capture)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)
    session.inspect_table("Data")
    session.build_snapshot("Data", "Team")
    session.shutdown()

    assert captures == [source]


def test_inspect_is_shallow_and_does_not_read_below_the_table(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)

    class MustNotRead:
        def __getattribute__(self, name: str):
            raise AssertionError(f"inspect scanned below Table via {name}")

    workbook.sheet.UsedRange = MustNotRead()
    excel = Excel(events, lambda _path: workbook)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)

    assert session.inspect_table("Data").columns == ("Team", "Amount")
    session.shutdown()


def test_opening_a_different_source_closes_the_previous_workbook(
    tmp_path: Path,
) -> None:
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    first.write_bytes(b"first")
    second.write_bytes(b"second")
    events: list[tuple[str, int]] = []
    first_workbook = Workbook(events, sheet_name="First")
    second_workbook = Workbook(events, sheet_name="Second")
    excel = Excel(
        events,
        lambda path: first_workbook if path == first else second_workbook,
    )
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()

    session.open_source(first)
    second_handle = session.open_source(second)

    assert second_handle.sheets == ("Second",)
    assert [name for name, _ in events].count("close") == 1
    assert len(excel.open_calls) == 2
    session.shutdown()
    assert [name for name, _ in events].count("close") == 2


def test_disk_change_closes_and_invalidates_the_open_source(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"before")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    excel = Excel(events, lambda _path: workbook)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)
    source.write_bytes(b"after")

    with pytest.raises(WorkbookValidationError, match="변경"):
        session.inspect_table("Data")
    with pytest.raises(RuntimeError, match="열려"):
        session.inspect_table("Data")

    assert [name for name, _ in events].count("close") == 1
    session.shutdown()


def test_worker_reraises_the_original_operation_exception(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    marker = LookupError("COM identity lost")

    class BrokenTables:
        @property
        def Count(self) -> int:
            raise marker

    workbook.sheet.ListObjects = BrokenTables()
    excel = Excel(events, lambda _path: workbook)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)

    with pytest.raises(LookupError) as caught:
        session.inspect_table("Data")

    assert caught.value is marker
    session.shutdown()


def test_shutdown_is_idempotent_and_orders_owned_resource_cleanup(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    excel = Excel(events, lambda _path: workbook)
    session = SourceSession(session_factory=session_factory_for(excel, events))

    session.start()
    session.start()
    session.open_source(source)
    session.shutdown()
    session.shutdown()

    assert [name for name, _ in events] == [
        "initialize",
        "close",
        "quit",
        "uninitialize",
    ]
    with pytest.raises(RuntimeError, match="종료"):
        session.start()
    with pytest.raises(RuntimeError, match="종료"):
        session.close_source()


def test_close_source_is_idempotent_before_shutdown(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    excel = Excel(events, lambda _path: workbook)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)

    session.close_source()
    session.close_source()
    session.shutdown()

    assert [name for name, _ in events].count("close") == 1


def test_save_plain_master_uses_safe_options_and_reopens_for_identity_check(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    events: list[tuple[str, int]] = []
    source_workbook = Workbook(events)
    verified_workbook = Workbook(events)
    save_calls: list[tuple[Path, dict[str, object], int]] = []

    def save_as(filename: str, **options: object) -> None:
        path = Path(filename)
        save_calls.append((path, options, threading.get_ident()))
        write_minimal_xlsx(path)

    source_workbook.SaveAs = save_as
    source_closed = False
    original_close = source_workbook.Close

    def close_source(**options: object) -> None:
        nonlocal source_closed
        source_closed = True
        original_close(**options)

    source_workbook.Close = close_source
    opened = iter((source_workbook, verified_workbook))

    def opener(_path: Path):
        workbook = next(opened)
        if workbook is verified_workbook:
            assert source_closed, "SaveAs workbook must close before master reopen"
        return workbook

    excel = Excel(events, opener)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)
    snapshot = session.build_snapshot("Data", "Team")
    try:
        master = session.save_plain_master(run_dir, snapshot)

        assert master.parent == run_dir
        assert master.name == "m.xlsx"
        assert master.is_file()
        assert save_calls[0][1] == {
            "FileFormat": 51,
            "Password": "",
            "WriteResPassword": "",
            "ReadOnlyRecommended": False,
            "AddToMru": False,
        }
        assert excel.open_calls[1][1]["Password"] == ""
        assert excel.open_calls[1][1]["WriteResPassword"] == ""
        assert [name for name, _ in events].count("close") == 2
    finally:
        session.shutdown()
    assert [name for name, _ in events].count("close") == 2


def test_snapshot_flags_artifacts_and_plain_master_cleans_them_after_save_as(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    events: list[tuple[str, int]] = []
    source_workbook = Workbook(events)
    source_workbook.sheet.Comments.Count = 1
    verified_workbook = Workbook(events)
    sequence: list[str] = []

    def save_as(filename: str, **_options: object) -> None:
        sequence.append("save-as")
        write_minimal_xlsx(Path(filename))

    source_workbook.SaveAs = save_as
    source_workbook.Save = lambda: sequence.append("save")

    def delete(sheet: object) -> None:
        assert sheet is source_workbook.sheet
        assert sequence == ["save-as"]
        sequence.append("delete")

    monkeypatch.setattr(
        "excel_splitter.source_session.delete_removable_artifacts", delete
    )
    opened = iter((source_workbook, verified_workbook))
    excel = Excel(events, lambda _path: next(opened))
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)
    snapshot = session.build_snapshot("Data", "Team")
    try:
        assert snapshot.has_removable_artifacts is True
        session.save_plain_master(run_dir, snapshot)
    finally:
        session.shutdown()

    assert sequence == ["save-as", "delete", "save"]


@pytest.mark.parametrize("failure", ["package", "identity"])
def test_save_plain_master_deletes_partial_output_on_verification_failure(
    tmp_path: Path, failure: str
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    events: list[tuple[str, int]] = []
    source_workbook = Workbook(events)
    verified_workbook = Workbook(
        events, table_name="Wrong" if failure == "identity" else "Orders"
    )

    def save_as(filename: str, **_options: object) -> None:
        if failure == "package":
            Path(filename).write_bytes(b"not a zip")
        else:
            write_minimal_xlsx(Path(filename))

    source_workbook.SaveAs = save_as
    opened = iter((source_workbook, verified_workbook))
    excel = Excel(events, lambda _path: next(opened))
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)
    snapshot = session.build_snapshot("Data", "Team")
    try:
        with pytest.raises(SplitExecutionError, match="master"):
            session.save_plain_master(run_dir, snapshot)

        assert tuple(run_dir.iterdir()) == ()
    finally:
        session.shutdown()


def test_save_plain_master_reports_plaintext_path_when_unlink_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    run_dir = tmp_path / "run"
    run_dir.mkdir()
    events: list[tuple[str, int]] = []
    workbook = Workbook(events)
    workbook.SaveAs = lambda filename, **_options: Path(filename).write_bytes(b"plain")
    excel = Excel(events, lambda _path: workbook)
    session = SourceSession(session_factory=session_factory_for(excel, events))
    session.start()
    session.open_source(source)
    snapshot = session.build_snapshot("Data", "Team")
    monkeypatch.setattr(
        Path, "unlink", lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("locked"))
    )
    try:
        with pytest.raises(SplitExecutionError, match="평문 master.*m.xlsx"):
            session.save_plain_master(run_dir, snapshot)
    finally:
        session.shutdown()
