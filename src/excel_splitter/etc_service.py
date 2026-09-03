from __future__ import annotations

import os
import shutil
import stat
import sys
import tempfile
from pathlib import Path

from .compare_service import _save_comparison as _save_xlsx, _validate_paths
from .errors import SplitExecutionError, WorkbookValidationError
from .excel_artifacts import delete_removable_artifacts
from .excel_gateway import _close_without_saving, _excel_session, _open_workbook, _publish_temp, _worksheet
from .file_signature import capture_signature, same_signature
from .ports import ProgressCallback
from .source_session import _verify_xlsx_package
from .split_service import _validate_source_path


def _subtract_rectangle(rectangle: tuple[int, int, int, int], excluded: tuple[int, int, int, int]) -> list[tuple[int, int, int, int]]:
    top, left, bottom, right = rectangle
    excluded_top, excluded_left, excluded_bottom, excluded_right = excluded
    overlap_top, overlap_left = max(top, excluded_top), max(left, excluded_left)
    overlap_bottom, overlap_right = min(bottom, excluded_bottom), min(right, excluded_right)
    if overlap_top > overlap_bottom or overlap_left > overlap_right:
        return [rectangle]
    return [candidate for candidate in (
        (top, left, overlap_top - 1, right), (overlap_bottom + 1, left, bottom, right),
        (overlap_top, left, overlap_bottom, overlap_left - 1), (overlap_top, overlap_right + 1, overlap_bottom, right),
    ) if candidate[0] <= candidate[2] and candidate[1] <= candidate[3]]


def _reset_fills_except_table_headers(sheet: object) -> None:
    headers = []
    for index in range(1, sheet.ListObjects.Count + 1):
        table = sheet.ListObjects.Item(index)
        if not table.ShowHeaders:
            continue
        header = table.HeaderRowRange
        if header is not None:
            headers.append((header.Row, header.Column, header.Row + header.Rows.Count - 1,
                            header.Column + header.Columns.Count - 1))
    if not headers:
        sheet.Cells.Interior.Pattern = -4142  # xlPatternNone; retain Table styles and conditional formats.
        return
    rectangles = [(1, 1, sheet.Rows.Count, sheet.Columns.Count)]
    for header in headers:
        rectangles = [remaining for rectangle in rectangles for remaining in _subtract_rectangle(rectangle, header)]
    for top, left, bottom, right in rectangles:
        sheet.Range(sheet.Cells.Item(top, left), sheet.Cells.Item(bottom, right)).Interior.Pattern = -4142


class EtcService:
    def inspect_source(self, source: Path) -> tuple[str, ...]:
        source = Path(source).resolve()
        _validate_source_path(source)
        signature = capture_signature(source)
        with _excel_session() as excel:
            workbook = _open_workbook(excel, source, read_only=True)
            try:
                sheets = tuple(str(workbook.Worksheets.Item(index).Name) for index in range(1, workbook.Worksheets.Count + 1))
            finally:
                _close_without_saving(workbook)
        if not same_signature(source, signature):
            raise WorkbookValidationError("시트 목록을 읽는 동안 원본이 변경되었습니다.")
        return sheets

    def execute(
        self, source: Path, sheet_name: str, target: Path, *,
        remove_artifacts: bool, reset_fill: bool, progress: ProgressCallback, exclude_table_headers: bool = True,
    ) -> Path:
        if not remove_artifacts and not reset_fill:
            raise WorkbookValidationError("실행할 정리 작업을 하나 이상 선택하세요.")
        source, target = Path(source).resolve(), Path(target).absolute()
        _validate_paths(source, source, target)
        target = target.resolve()
        signature = capture_signature(source)
        descriptor, filename = tempfile.mkstemp(prefix=".et-", suffix=".xlsx", dir=target.parent)
        os.close(descriptor)
        temp = Path(filename)
        try:
            shutil.copy2(source, temp)
            temp.chmod(stat.S_IREAD | stat.S_IWRITE)
            if capture_signature(temp) != signature:
                raise WorkbookValidationError("원본 파일을 복사하는 동안 변경되었습니다.")
            with _excel_session() as excel:
                workbook = _open_workbook(excel, temp, read_only=False)
                try:
                    sheet = _worksheet(workbook, sheet_name)
                    if sheet.ProtectContents or (remove_artifacts and sheet.ProtectDrawingObjects):
                        raise WorkbookValidationError(f"보호된 시트는 정리할 수 없습니다: {sheet_name}")
                    if remove_artifacts:
                        delete_removable_artifacts(sheet)
                    if reset_fill:
                        if exclude_table_headers:
                            _reset_fills_except_table_headers(sheet)
                        else:
                            sheet.Cells.Interior.Pattern = -4142  # xlPatternNone; retain Table styles and conditional formats.
                    progress(1, 1, sheet_name)
                    _save_xlsx(workbook, temp)
                finally:
                    _close_without_saving(workbook)
            _verify_xlsx_package(temp)
            _validate_paths(source, source, target)
            if not same_signature(source, signature):
                raise WorkbookValidationError("정리하는 동안 원본이 변경되었습니다.")
            _publish_temp(temp, target, None)
            return target
        finally:
            active_error = sys.exception()
            try:
                temp.unlink(missing_ok=True)
            except OSError as exc:
                raise SplitExecutionError(f"정리 임시 파일을 삭제하지 못했습니다: {temp}; {active_error or exc}") from exc
