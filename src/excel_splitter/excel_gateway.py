from __future__ import annotations

import os
import shutil
import sys
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
    SplitFailure,
    SplitResult,
    TableInfo,
    WorkbookSnapshot,
)
from .ports import ProgressCallback


_READ_OPEN_OPTIONS = {
    "UpdateLinks": 0,
    "ReadOnly": True,
    "AddToMru": False,
    "Notify": False,
    "IgnoreReadOnlyRecommended": True,
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
    if getattr(cell, "CommentThreaded", None) is not None:
        return True
    style = getattr(cell, "Style", "Normal")
    return style not in (None, "Normal")


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
    if any(_cell_has_unsupported_content(cell) for cell in _iter_cells(bounded)):
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
    def list_worksheets(self, source: Path) -> tuple[str, ...]:
        with _excel_session() as excel:
            workbook = _open_workbook(excel, source, read_only=True)
            try:
                return tuple(
                    str(workbook.Worksheets.Item(index).Name)
                    for index in range(1, workbook.Worksheets.Count + 1)
                )
            finally:
                _close_without_saving(workbook)

    def inspect_table(self, source: Path, sheet_name: str) -> TableInfo:
        with _excel_session() as excel:
            workbook = _open_workbook(excel, source, read_only=True)
            try:
                _sheet, table = _validated_table(workbook, sheet_name)
                columns = tuple(
                    str(table.ListColumns.Item(index).Name)
                    for index in range(1, table.ListColumns.Count + 1)
                )
                return TableInfo(
                    sheet_name=str(sheet_name),
                    table_name=str(table.Name),
                    columns=columns,
                    row_count=int(table.ListRows.Count),
                )
            finally:
                _close_without_saving(workbook)

    def build_snapshot(
        self, source: Path, sheet_name: str, column_name: str
    ) -> WorkbookSnapshot:
        signature = _capture_signature(source)
        with _excel_session() as excel:
            workbook = _open_workbook(excel, source, read_only=True)
            try:
                _sheet, table = _validated_table(workbook, sheet_name)
                column_index = _column_index(table, column_name)
                samples: list[CellSample] = []
                for row_index in range(1, table.ListRows.Count + 1):
                    cell = table.ListRows.Item(row_index).Range.Cells.Item(
                        1, column_index
                    )
                    value = cell.Value2
                    error_code = _excel_error_code(value)
                    samples.append(
                        CellSample(
                            row_index=row_index,
                            value=value,
                            text=str(cell.Text),
                            is_error=error_code is not None,
                            error_code=error_code,
                        )
                    )
                table_name = str(table.Name)
                row_count = int(table.ListRows.Count)
            finally:
                _close_without_saving(workbook)
        if _capture_signature(source) != signature:
            raise WorkbookValidationError("원본이 검사 중 변경되어 다시 검사해야 합니다.")
        groups = tuple(_group_samples(samples))
        return WorkbookSnapshot(
            source=source,
            signature=signature,
            sheet_name=sheet_name,
            table_name=table_name,
            column_name=column_name,
            row_count=row_count,
            groups=groups,
        )

    def write_groups(
        self,
        snapshot: WorkbookSnapshot,
        targets: tuple[OutputTarget, ...],
        progress: ProgressCallback,
    ) -> SplitResult:
        succeeded: list[Path] = []
        failed: list[SplitFailure] = []
        if not targets:
            return SplitResult((), ())
        target_parent = targets[0].path.parent
        normalized_parent = target_parent.resolve(strict=False)
        if any(
            target.path.parent.resolve(strict=False) != normalized_parent
            for target in targets[1:]
        ):
            raise SplitExecutionError("모든 결과 파일은 같은 출력 폴더에 있어야 합니다.")
        master = _copy_to_master(
            snapshot.source, snapshot.signature, target_parent
        )
        try:
            with _excel_session() as excel:
                for completed, target in enumerate(targets, start=1):
                    progress(completed - 1, len(targets), target.label)
                    try:
                        output = _write_one_group(excel, master, snapshot, target)
                        succeeded.append(output)
                    except OSError as exc:
                        failed.append(SplitFailure(target.label, str(exc)))
                    progress(completed, len(targets), target.label)
        finally:
            active_error = sys.exception()
            try:
                master.unlink(missing_ok=True)
            except OSError as exc:
                if active_error is None:
                    raise SplitExecutionError(
                        f"master 임시 파일을 삭제하지 못했습니다: {exc}"
                    ) from exc
        return SplitResult(tuple(succeeded), tuple(failed))


def _write_one_group(
    excel: Any,
    master: Path,
    snapshot: WorkbookSnapshot,
    target: OutputTarget,
) -> Path:
    temp_path = target.path.parent / f".{target.path.stem}.{uuid.uuid4().hex}.xlsx"
    workbook = None
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
        _verify_target_unchanged(target.path, target.prior_signature)
        os.replace(temp_path, target.path)
        return target.path
    finally:
        active_error = sys.exception()
        close_error: Exception | None = None
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception as exc:
                close_error = exc
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
