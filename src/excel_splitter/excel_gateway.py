from __future__ import annotations

import os
import logging
import shutil
import sys
import tempfile
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from .errors import (
    ExcelSplitterError,
    ExcelUnavailableError,
    SplitExecutionError,
    WorkbookValidationError,
)
from .models import (
    CellSample,
    FileSignature,
    OutputTarget,
    SplitResult,
    TableInfo,
    WorkbookSnapshot,
)
from .ports import ProgressCallback
from .parallel_writer import write_targets


_READ_OPEN_OPTIONS = {
    "UpdateLinks": 0,
    "ReadOnly": True,
    "AddToMru": False,
    "Notify": False,
    "IgnoreReadOnlyRecommended": True,
    "Password": "",
    "WriteResPassword": "",
}
_WRITE_OPEN_OPTIONS = {**_READ_OPEN_OPTIONS, "ReadOnly": False}
_APPLICATION_SETTINGS = (
    "Visible",
    "DisplayAlerts",
    "EnableEvents",
    "ScreenUpdating",
    "AskToUpdateLinks",
    "AutomationSecurity",
)
_SAFE_APPLICATION_VALUES = (False, False, False, False, False, 3)

# CVErr values returned by Excel through Value2 (signed HRESULT form).
_EXCEL_ERROR_CODES = {2000, 2007, 2015, 2023, 2029, 2036, 2042, 2043, 2045}


def _capture_signature(path: Path) -> FileSignature:
    from .file_signature import capture_signature

    return capture_signature(path)


def _group_samples(samples: list[CellSample]):
    from .classifier import group_samples

    return group_samples(samples)


def _same_signature(path: Path, expected: FileSignature | None) -> bool:
    from .file_signature import same_signature

    return same_signature(path, expected)


@contextmanager
def _excel_session() -> Iterator[Any]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ExcelUnavailableError("pywin32를 불러올 수 없습니다.") from exc

    pythoncom.CoInitialize()
    excel = None
    original_settings: dict[str, Any] = {}
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        for name, safe_value in zip(
            _APPLICATION_SETTINGS, _SAFE_APPLICATION_VALUES, strict=True
        ):
            original_settings[name] = getattr(excel, name, None)
            setattr(excel, name, safe_value)
        yield excel
    except ExcelSplitterError:
        raise
    except AttributeError:
        threaded_comment = None
    except Exception as exc:
        raise SplitExecutionError(f"Excel 자동화에 실패했습니다: {exc}") from exc
    finally:
        if excel is not None:
            for name, value in original_settings.items():
                try:
                    setattr(excel, name, value)
                except Exception:
                    pass
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _single_table(tables: Any) -> Any:
    if tables.Count == 0:
        raise WorkbookValidationError("정식 Excel Table이 없습니다.")
    if tables.Count > 1:
        raise WorkbookValidationError("이 시트에는 Table이 2개 이상 있습니다.")
    return tables.Item(1)


def _delete_rows(rows: Any, indexes: tuple[int, ...] | list[int]) -> None:
    for index in sorted(indexes, reverse=True):
        rows(index).Delete()


def _remove_other_sheets(sheets: Any, selected_name: str) -> None:
    for index in range(sheets.Count, 0, -1):
        sheet = sheets.Item(index)
        if sheet.Name != selected_name:
            sheet.Delete()


def _column_index(table: Any, column_name: str) -> int:
    matches = [
        index
        for index in range(1, table.ListColumns.Count + 1)
        if table.ListColumns.Item(index).Name == column_name
    ]
    if len(matches) != 1:
        raise WorkbookValidationError(
            "선택한 컬럼 이름과 정확히 하나 일치하는 Table 열이 필요합니다."
        )
    return matches[0]


def _worksheet(workbook: Any, sheet_name: str) -> Any:
    matches = [
        workbook.Worksheets.Item(index)
        for index in range(1, workbook.Worksheets.Count + 1)
        if workbook.Worksheets.Item(index).Name == sheet_name
    ]
    if len(matches) != 1:
        raise WorkbookValidationError("선택한 워크시트를 정확히 찾을 수 없습니다.")
    return matches[0]


def _iter_cells(cell_range: Any) -> Iterator[Any]:
    cells = cell_range.Cells
    for index in range(1, cells.Count + 1):
        yield cells.Item(index)


def _cell_has_unsupported_content(cell: Any) -> bool:
    if cell.Value2 not in (None, ""):
        return True
    formula = getattr(cell, "Formula", None)
    if formula not in (None, ""):
        return True
    if bool(getattr(cell, "MergeCells", False)):
        return True
    if getattr(getattr(cell, "Hyperlinks", None), "Count", 0):
        return True
    if getattr(cell, "Comment", None) is not None:
        return True
    try:
        threaded_comment = getattr(cell, "CommentThreaded", None)
    except Exception as exc:
        if not _unsupported_threaded_comments_error(exc):
            raise
        threaded_comment = None
    if threaded_comment is not None:
        return True
    style = getattr(cell, "Style", "Normal")
    return style not in (None, "Normal")


def _bulk_contains_value(value: Any) -> bool:
    if isinstance(value, (tuple, list)):
        return any(_bulk_contains_value(item) for item in value)
    return value not in (None, "")


def _unsupported_threaded_comments_error(exc: Exception) -> bool:
    hresult = getattr(exc, "hresult", exc.args[0] if exc.args else None)
    if exc.__class__.__name__ != "com_error":
        return False
    if hresult in (-2147352573, -2147352570):
        return True
    message = str(exc).casefold()
    return hresult == -2147352567 and (
        "commentthreaded" in message
        or "threaded comment" in message
        or "스레드 주석" in message
    )


def _threaded_comments_count(owner: Any) -> int:
    try:
        collection = getattr(owner, "CommentsThreaded")
    except AttributeError:
        return 0
    except Exception as exc:
        if _unsupported_threaded_comments_error(exc):
            return 0
        raise
    try:
        return int(collection.Count)
    except Exception as exc:
        if _unsupported_threaded_comments_error(exc):
            return 0
        raise


def _validate_below_table(sheet: Any, table: Any) -> None:
    first_row = table.Range.Row + table.Range.Rows.Count
    used_last_row = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1
    if first_row > used_last_row:
        return
    first_column = table.Range.Column
    last_column = first_column + table.Range.Columns.Count - 1
    bounded = sheet.Range(
        sheet.Cells(first_row, first_column),
        sheet.Cells(used_last_row, last_column),
    )
    values = bounded.Value2
    formulas = bounded.Formula
    merge_cells = bounded.MergeCells
    has_content = (
        _bulk_contains_value(values)
        or _bulk_contains_value(formulas)
        or merge_cells not in (False, None)
        or int(bounded.Hyperlinks.Count) > 0
    )
    if has_content:
        raise WorkbookValidationError(
            "Table 아래 열 범위에 삭제 시 이동될 수 있는 콘텐츠가 있습니다."
        )

    style = bounded.Style
    if style not in (None, "Normal"):
        raise WorkbookValidationError(
            "Table 아래 열 범위에 삭제 시 이동될 수 있는 콘텐츠가 있습니다."
        )
    needs_cell_scan = (
        merge_cells is None
        or style is None
        or int(sheet.Comments.Count) > 0
        or _threaded_comments_count(sheet) > 0
    )
    if needs_cell_scan and any(
        _cell_has_unsupported_content(cell) for cell in _iter_cells(bounded)
    ):
        raise WorkbookValidationError(
            "Table 아래 열 범위에 삭제 시 이동될 수 있는 콘텐츠가 있습니다."
        )


def _validated_table(workbook: Any, sheet_name: str) -> tuple[Any, Any]:
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
    _validate_below_table(sheet, table)
    return sheet, table


def _open_workbook(excel: Any, source: Path, *, read_only: bool) -> Any:
    options = _READ_OPEN_OPTIONS if read_only else _WRITE_OPEN_OPTIONS
    return excel.Workbooks.Open(str(source), **options)


def _close_without_saving(workbook: Any) -> None:
    workbook.Close(SaveChanges=False)


def _excel_error_code(value: Any) -> int | None:
    if type(value) is int:
        code = value & 0xFFFF if value < 0 else value
        if code in _EXCEL_ERROR_CODES:
            return code
    return None


def _verify_target_unchanged(
    path: Path, prior_signature: FileSignature | None
) -> None:
    if prior_signature is None:
        if path.exists():
            raise OSError(f"대상 파일이 새로 생성되었습니다: {path}")
        return
    if not path.is_file():
        raise OSError(f"기존 대상 파일이 없어졌습니다: {path}")
    if not _same_signature(path, prior_signature):
        raise OSError(f"대상 파일이 미리보기 후 변경되었습니다: {path}")


def _restore_recovery(backup: Path, target: Path, reason: str) -> OSError:
    if not target.exists():
        try:
            os.rename(backup, target)
        except OSError:
            pass
    if backup.exists():
        return OSError(f"{reason} 복구 파일을 보존했습니다: {backup}")
    return OSError(reason)


def _publish_temp(
    temp_path: Path,
    target_path: Path,
    prior_signature: FileSignature | None,
) -> None:
    if prior_signature is None:
        _verify_target_unchanged(target_path, None)
        os.rename(temp_path, target_path)
        return

    backup = target_path.parent / f".esr-{uuid.uuid4().hex[:12]}.xlsx"
    try:
        os.rename(target_path, backup)
    except OSError as exc:
        raise OSError(f"기존 대상 파일을 점유하지 못했습니다: {target_path}: {exc}") from exc

    try:
        backup_signature = _capture_signature(backup)
    except OSError as exc:
        raise _restore_recovery(
            backup, target_path, f"점유한 대상 파일을 검증하지 못했습니다: {exc}"
        ) from exc
    if backup_signature != prior_signature:
        raise _restore_recovery(
            backup, target_path, "대상 파일이 미리보기 후 변경되었습니다."
        )
    try:
        os.rename(temp_path, target_path)
    except OSError as exc:
        raise _restore_recovery(
            backup, target_path, f"결과 파일을 게시하지 못했습니다: {exc}"
        ) from exc
    try:
        backup.unlink()
    except OSError as exc:
        raise OSError(
            f"결과는 게시했지만 복구 파일을 삭제하지 못했습니다: {backup}: {exc}"
        ) from exc


def _copy_to_master(
    source: Path, expected: FileSignature, master_parent: Path
) -> Path:
    before = _capture_signature(source)
    if before != expected:
        raise SplitExecutionError("원본이 미리보기 후 변경되었습니다.")
    master = master_parent / f".{source.stem}.master.{uuid.uuid4().hex}.xlsx"
    try:
        shutil.copy2(source, master)
        after = _capture_signature(source)
        copied = _capture_signature(master)
        if after != expected or copied != expected:
            raise SplitExecutionError("복사된 master의 파일 서명이 원본과 다릅니다.")
        return master
    except Exception:
        master.unlink(missing_ok=True)
        raise


def _shape_snapshot(sheet: Any) -> tuple[tuple[str, float, float, float, float], ...]:
    shapes = sheet.Shapes
    return tuple(
        (
            str(shapes.Item(index).Name),
            shapes.Item(index).Left,
            shapes.Item(index).Top,
            shapes.Item(index).Width,
            shapes.Item(index).Height,
        )
        for index in range(1, shapes.Count + 1)
    )


def _restore_shapes(sheet: Any, snapshot: tuple[tuple[str, float, float, float, float], ...]) -> None:
    for name, left, top, width, height in snapshot:
        shape = sheet.Shapes.Item(name)
        shape.Left = left
        shape.Top = top
        shape.Width = width
        shape.Height = height


class ExcelComGateway:
    def __init__(self, source_session: Any | None = None) -> None:
        if source_session is None:
            # Local import avoids the helper dependency from source_session back
            # into this module while keeping one persistent owner per gateway.
            from .source_session import SourceSession

            source_session = SourceSession()
        self._source_session = source_session
        self._started = False
        self._active_source: Path | None = None
        self._logger = logging.getLogger("excel_splitter")

    def _ensure_started(self) -> None:
        if not self._started:
            self._source_session.start()
            self._started = True

    def prewarm(self) -> None:
        self._ensure_started()

    def _open_source(self, source: Path) -> Any:
        self._active_source = None
        try:
            handle = self._source_session.open_source(source)
        except Exception:
            try:
                self._source_session.close_source()
            except Exception:
                pass
            raise
        self._active_source = Path(source)
        return handle

    def _discard_source_after_error(self, error: Exception) -> None:
        self._active_source = None
        if isinstance(error, WorkbookValidationError):
            return
        try:
            self._source_session.close_source()
        except Exception:
            pass

    def list_worksheets(self, source: Path) -> tuple[str, ...]:
        self._ensure_started()
        started = time.perf_counter()
        handle = self._open_source(source)
        self._logger.info(
            "operation=list_worksheets elapsed_seconds=%.3f sheet_count=%d",
            time.perf_counter() - started,
            len(handle.sheets),
        )
        return handle.sheets

    def inspect_table(self, source: Path, sheet_name: str) -> TableInfo:
        self._ensure_started()
        started = time.perf_counter()
        if self._active_source != Path(source):
            self._open_source(source)
        try:
            info = self._source_session.inspect_table(sheet_name)
        except Exception as exc:
            self._discard_source_after_error(exc)
            raise
        self._logger.info(
            "operation=inspect_table elapsed_seconds=%.3f row_count=%d column_count=%d",
            time.perf_counter() - started,
            info.row_count,
            len(info.columns),
        )
        return info

    def build_snapshot(
        self, source: Path, sheet_name: str, column_name: str
    ) -> WorkbookSnapshot:
        self._ensure_started()
        started = time.perf_counter()
        if self._active_source != Path(source):
            self._open_source(source)
        try:
            snapshot = self._source_session.build_snapshot(sheet_name, column_name)
        except Exception as exc:
            self._discard_source_after_error(exc)
            raise
        self._logger.info(
            "operation=build_snapshot elapsed_seconds=%.3f row_count=%d group_count=%d",
            time.perf_counter() - started,
            snapshot.row_count,
            len(snapshot.groups),
        )
        return snapshot

    def write_groups(
        self,
        snapshot: WorkbookSnapshot,
        targets: tuple[OutputTarget, ...],
        progress: ProgressCallback,
    ) -> SplitResult:
        started = time.perf_counter()
        target_parent = targets[0].path.parent if targets else snapshot.source.parent
        normalized_parent = target_parent.resolve(strict=False)
        if any(
            target.path.parent.resolve(strict=False) != normalized_parent
            for target in targets[1:]
        ):
            raise SplitExecutionError("모든 결과 파일은 같은 출력 폴더에 있어야 합니다.")
        self._ensure_started()
        run_dir = Path(tempfile.mkdtemp(prefix=".e", dir=target_parent))
        try:
            longest_excel_temp = run_dir / "g-000000000000.xlsx"
            if len(str(longest_excel_temp.resolve(strict=False))) > 218:
                raise SplitExecutionError(
                    "출력 폴더가 너무 깊어 Excel 임시 경로가 218자를 넘습니다."
                )
            if self._active_source != snapshot.source:
                self._open_source(snapshot.source)
            try:
                master = self._source_session.save_plain_master(run_dir, snapshot)
            finally:
                self._source_session.close_source()
                self._active_source = None
            result = write_targets(
                master,
                snapshot,
                targets,
                _write_one_group,
                _excel_session,
                progress,
            )
            self._logger.info(
                "operation=write_groups elapsed_seconds=%.3f target_count=%d success_count=%d failure_count=%d",
                time.perf_counter() - started,
                len(targets),
                len(result.succeeded),
                len(result.failed),
            )
            return result
        finally:
            active_error = sys.exception()
            cleanup_error: OSError | None = None
            for attempt in range(3):
                try:
                    shutil.rmtree(run_dir)
                    cleanup_error = None
                    break
                except OSError as exc:
                    cleanup_error = exc
                    if attempt < 2:
                        time.sleep(0.05)
            if cleanup_error is not None:
                message = f"평문 임시 폴더가 남았습니다: {run_dir}: {cleanup_error}"
                if active_error is not None:
                    message = f"{active_error}; {message}"
                    raise SplitExecutionError(message) from active_error
                raise SplitExecutionError(message) from cleanup_error

    def shutdown(self) -> None:
        try:
            self._source_session.shutdown()
        finally:
            self._active_source = None


def _write_one_group(
    excel: Any,
    master: Path,
    snapshot: WorkbookSnapshot,
    target: OutputTarget,
) -> Path:
    temp_path = master.parent / f"g-{uuid.uuid4().hex[:12]}.xlsx"
    workbook = None
    sheet = None
    table = None
    try:
        shutil.copy2(master, temp_path)
        workbook = _open_workbook(excel, temp_path, read_only=False)
        sheet, table = _validated_table(workbook, snapshot.sheet_name)
        if str(table.Name) != snapshot.table_name:
            raise SplitExecutionError("Table 식별자가 미리보기와 다릅니다.")
        _column_index(table, snapshot.column_name)
        if int(table.ListRows.Count) != snapshot.row_count:
            raise SplitExecutionError("Table 행 수가 미리보기와 다릅니다.")

        shapes = _shape_snapshot(sheet)
        if bool(getattr(sheet, "FilterMode", False)):
            sheet.ShowAllData()
        selected = next(
            (group for group in snapshot.groups if group.key == target.key), None
        )
        if selected is None:
            raise SplitExecutionError("출력 대상 분류를 스냅샷에서 찾을 수 없습니다.")
        retained = set(selected.row_indexes)
        remove = tuple(
            index for index in range(1, snapshot.row_count + 1) if index not in retained
        )
        _delete_rows(table.ListRows, remove)
        _remove_other_sheets(workbook.Sheets, snapshot.sheet_name)
        if workbook.Sheets.Count != 1:
            raise SplitExecutionError("선택 워크시트 하나만 남기지 못했습니다.")
        if int(table.ListRows.Count) != len(selected.row_indexes):
            raise SplitExecutionError("결과 Table 행 수가 예상과 다릅니다.")
        _restore_shapes(sheet, shapes)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        _publish_temp(temp_path, target.path, target.prior_signature)
        return target.path
    finally:
        active_error = sys.exception()
        close_error: Exception | None = None
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception as exc:
                close_error = exc
        table = None
        sheet = None
        workbook = None
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            if active_error is None and close_error is None:
                raise
        if close_error is not None and (
            active_error is None or isinstance(active_error, OSError)
        ):
            raise SplitExecutionError(
                f"결과 통합문서를 닫지 못했습니다: {close_error}"
            ) from close_error
