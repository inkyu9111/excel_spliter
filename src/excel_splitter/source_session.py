from __future__ import annotations

import queue
import threading
import zipfile
from concurrent.futures import Future
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, ContextManager, TypeVar

from .errors import SplitExecutionError, WorkbookValidationError
from .excel_artifacts import delete_removable_artifacts, has_removable_artifacts
from .excel_gateway import (
    _close_without_saving,
    _column_index,
    _excel_session,
    _open_workbook,
    _single_table,
    _validate_below_table,
    _worksheet,
)
from .file_signature import capture_signature
from .models import FileSignature, TableInfo, WorkbookSnapshot


_Result = TypeVar("_Result")
_STOP = object()
_REQUIRED_PACKAGE_PARTS = (
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
)


@dataclass(frozen=True)
class SourceHandleInfo:
    source: Path
    signature: FileSignature
    sheets: tuple[str, ...]


def _group_bulk_samples(data_body_range: object, row_count: int):
    from .snapshot_reader import group_bulk_samples

    return group_bulk_samples(data_body_range, row_count)


def _shallow_validated_table(workbook: Any, sheet_name: str) -> tuple[Any, Any]:
    if bool(getattr(workbook, "ProtectStructure", False)):
        raise WorkbookValidationError("통합문서 구조가 보호되어 있습니다.")
    sheet = _worksheet(workbook, sheet_name)
    if getattr(sheet, "Visible", -1) != -1:
        raise WorkbookValidationError("숨겨진 워크시트는 분할할 수 없습니다.")
    if bool(getattr(sheet, "ProtectContents", False)):
        raise WorkbookValidationError("선택한 워크시트가 보호되어 있습니다.")
    table = _single_table(sheet.ListObjects)
    if table.SourceType != 1:
        raise WorkbookValidationError("외부 연결 Table은 지원하지 않습니다.")
    if table.ListRows.Count == 0:
        raise WorkbookValidationError("Table에 데이터 행이 없습니다.")
    return sheet, table


def _table_info(sheet_name: str, sheet: Any, table: Any) -> TableInfo:
    columns = tuple(
        str(table.ListColumns.Item(index).Name)
        for index in range(1, table.ListColumns.Count + 1)
    )
    return TableInfo(
        sheet_name=sheet_name,
        table_name=str(table.Name),
        columns=columns,
        row_count=int(table.ListRows.Count),
    )


def _verify_xlsx_package(path: Path) -> None:
    try:
        with zipfile.ZipFile(path) as package:
            if package.testzip() is not None:
                raise ValueError("CRC 검사가 실패했습니다.")
            for name in _REQUIRED_PACKAGE_PARTS:
                package.read(name)
    except (OSError, KeyError, ValueError, zipfile.BadZipFile) as exc:
        raise SplitExecutionError(
            f"비보호 master의 Excel package를 검증하지 못했습니다: {exc}"
        ) from exc


class SourceSession:
    def __init__(
        self,
        *,
        session_factory: Callable[[], ContextManager[Any]] | None = None,
    ) -> None:
        self._session_factory = session_factory or _excel_session
        self._requests: queue.Queue[object] = queue.Queue()
        self._state_lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._started = False
        self._shutdown = False
        self._startup: Future[None] | None = None

        self._excel: Any = None
        self._workbook: Any = None
        self._handle: SourceHandleInfo | None = None

    def start(self) -> None:
        with self._state_lock:
            if self._shutdown:
                raise RuntimeError("source session이 이미 종료되었습니다.")
            if self._started:
                startup = self._startup
            else:
                self._started = True
                startup = Future()
                self._startup = startup
                self._thread = threading.Thread(
                    target=self._worker_main,
                    name="excel-source-session",
                    daemon=False,
                )
                self._thread.start()
        assert startup is not None
        startup.result()

    def open_source(self, source: Path) -> SourceHandleInfo:
        return self._call(self._open_source, Path(source))

    def inspect_table(self, sheet_name: str) -> TableInfo:
        return self._call(self._inspect_table, sheet_name)

    def build_snapshot(
        self, sheet_name: str, column_name: str
    ) -> WorkbookSnapshot:
        return self._call(self._build_snapshot, sheet_name, column_name)

    def save_plain_master(
        self, run_dir: Path, snapshot: WorkbookSnapshot
    ) -> Path:
        return self._call(self._save_plain_master, Path(run_dir), snapshot)

    def close_source(self) -> None:
        self._call(self._close_source)

    def shutdown(self) -> None:
        with self._state_lock:
            if self._shutdown:
                return
            self._shutdown = True
            if not self._started:
                return
            self._requests.put(_STOP)
            thread = self._thread
        assert thread is not None
        thread.join()

    def _call(self, operation: Callable[..., _Result], *args: object) -> _Result:
        with self._state_lock:
            if self._shutdown:
                raise RuntimeError("source session이 이미 종료되었습니다.")
            if not self._started:
                raise RuntimeError("source session을 먼저 시작해야 합니다.")
            future: Future[_Result] = Future()
            self._requests.put((operation, args, future))
        return future.result()

    def _worker_main(self) -> None:
        assert self._startup is not None
        try:
            with self._session_factory() as excel:
                self._excel = excel
                self._startup.set_result(None)
                while True:
                    request = self._requests.get()
                    if request is _STOP:
                        break
                    operation, args, future = request
                    if not future.set_running_or_notify_cancel():
                        continue
                    try:
                        future.set_result(operation(*args))
                    except BaseException as exc:
                        future.set_exception(exc)
                self._close_source()
        except BaseException as exc:
            if not self._startup.done():
                self._startup.set_exception(exc)
        finally:
            self._excel = None

    def _open_source(self, source: Path) -> SourceHandleInfo:
        signature = capture_signature(source)
        if self._handle is not None:
            if self._handle.source == source and self._handle.signature == signature:
                return self._handle
            self._close_source()
        workbook = None
        try:
            workbook = _open_workbook(self._excel, source, read_only=True)
            sheets = tuple(
                str(workbook.Worksheets.Item(index).Name)
                for index in range(1, workbook.Worksheets.Count + 1)
            )
            handle = SourceHandleInfo(source, signature, sheets)
            self._workbook = workbook
            self._handle = handle
            return handle
        except BaseException:
            if workbook is not None:
                _close_without_saving(workbook)
            raise

    def _require_open_source(self) -> SourceHandleInfo:
        if self._workbook is None or self._handle is None:
            raise RuntimeError("열려 있는 source 통합문서가 없습니다.")
        return self._handle

    def _ensure_source_unchanged(self, *, for_save: bool = False) -> SourceHandleInfo:
        handle = self._require_open_source()
        try:
            if for_save:
                unchanged = capture_signature(handle.source) == handle.signature
            else:
                stat = handle.source.stat()
                unchanged = (stat.st_size, stat.st_mtime_ns) == (
                    handle.signature.size,
                    handle.signature.mtime_ns,
                )
        except OSError:
            self._close_source()
            raise
        if not unchanged:
            self._close_source()
            error_type = SplitExecutionError if for_save else WorkbookValidationError
            raise error_type("원본이 열린 후 변경되어 다시 선택해야 합니다.")
        return handle

    def _inspect_table(self, sheet_name: str) -> TableInfo:
        self._ensure_source_unchanged()
        sheet, table = _shallow_validated_table(self._workbook, sheet_name)
        return _table_info(sheet_name, sheet, table)

    def _build_snapshot(
        self, sheet_name: str, column_name: str
    ) -> WorkbookSnapshot:
        handle = self._ensure_source_unchanged()
        sheet, table = _shallow_validated_table(self._workbook, sheet_name)
        _validate_below_table(sheet, table)
        column_index = _column_index(table, column_name)
        row_count = int(table.ListRows.Count)
        groups = tuple(
            _group_bulk_samples(
                table.ListColumns.Item(column_index).DataBodyRange,
                row_count,
            )
        )
        return WorkbookSnapshot(
            source=handle.source,
            signature=handle.signature,
            sheet_name=sheet_name,
            table_name=str(table.Name),
            column_name=column_name,
            row_count=row_count,
            groups=groups,
            has_removable_artifacts=has_removable_artifacts(sheet),
        )

    def _save_plain_master(
        self, run_dir: Path, snapshot: WorkbookSnapshot
    ) -> Path:
        handle = self._ensure_source_unchanged(for_save=True)
        if snapshot.source != handle.source or snapshot.signature != handle.signature:
            raise SplitExecutionError("미리보기와 열린 원본이 일치하지 않습니다.")
        if not run_dir.is_dir():
            raise SplitExecutionError("master 실행 폴더가 존재하지 않습니다.")
        master = run_dir / "m.xlsx"
        verified_workbook = None
        try:
            self._workbook.SaveAs(
                str(master),
                FileFormat=51,
                Password="",
                WriteResPassword="",
                ReadOnlyRecommended=False,
                AddToMru=False,
            )
            try:
                open_master_key = str(
                    Path(str(self._workbook.FullName)).resolve(strict=False)
                ).casefold()
                intended_master_key = str(master.resolve(strict=False)).casefold()
            except Exception as exc:
                raise SplitExecutionError(
                    f"SaveAs 후 master 경로를 확인할 수 없습니다: {exc}"
                ) from exc
            if open_master_key != intended_master_key:
                raise SplitExecutionError(
                    "SaveAs 후 열린 통합문서가 master 경로와 일치하지 않습니다."
                )
            if snapshot.has_removable_artifacts:
                sheet = _worksheet(self._workbook, snapshot.sheet_name)
                delete_removable_artifacts(sheet)
                self._workbook.Save()
            # Validate that in-memory copy first, then release it before the
            # required reopen check; Excel may reject opening the same path
            # while the SaveAs workbook still owns it.
            _sheet, table = _shallow_validated_table(
                self._workbook, snapshot.sheet_name
            )
            if str(table.Name) != snapshot.table_name:
                raise SplitExecutionError("master의 Table 식별자가 다릅니다.")
            _column_index(table, snapshot.column_name)
            if int(table.ListRows.Count) != snapshot.row_count:
                raise SplitExecutionError("master의 Table 행 수가 다릅니다.")
            _verify_xlsx_package(master)
            self._close_source()
            verified_workbook = _open_workbook(
                self._excel, master, read_only=True
            )
            _sheet, table = _shallow_validated_table(
                verified_workbook, snapshot.sheet_name
            )
            if str(table.Name) != snapshot.table_name:
                raise SplitExecutionError("master의 Table 식별자가 다릅니다.")
            _column_index(table, snapshot.column_name)
            if int(table.ListRows.Count) != snapshot.row_count:
                raise SplitExecutionError("master의 Table 행 수가 다릅니다.")
            _close_without_saving(verified_workbook)
            verified_workbook = None
            return master
        except BaseException as exc:
            if verified_workbook is not None:
                try:
                    _close_without_saving(verified_workbook)
                except Exception:
                    pass
            try:
                self._close_source()
            except Exception:
                pass
            cleanup_error: OSError | None = None
            try:
                master.unlink(missing_ok=True)
            except OSError as unlink_error:
                cleanup_error = unlink_error
            if cleanup_error is not None:
                raise SplitExecutionError(
                    f"{exc}; 평문 master가 남았습니다: {master}: {cleanup_error}"
                ) from exc
            if isinstance(exc, SplitExecutionError):
                raise
            raise SplitExecutionError(
                f"비보호 master를 만들 수 없습니다: {exc}"
            ) from exc

    def _close_source(self) -> None:
        workbook = self._workbook
        self._workbook = None
        self._handle = None
        if workbook is not None:
            _close_without_saving(workbook)
