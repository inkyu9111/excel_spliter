from __future__ import annotations

import os
import shutil
import stat
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import SplitExecutionError, WorkbookValidationError
from .excel_artifacts import unsupported_threaded_comments_error
from .excel_gateway import (
    _bulk_contains_value,
    _close_without_saving,
    _excel_session,
    _open_workbook,
    _publish_temp,
    _single_table,
)
from .file_signature import capture_signature, same_signature
from .models import FileSignature
from .naming import _INVALID_CHARACTERS
from .ports import ProgressCallback
from .split_service import _validate_output_dir, _validate_source_path
from .source_session import _verify_xlsx_package


@dataclass(frozen=True)
class MergeInput:
    source: Path
    signature: FileSignature
    sheet_name: str
    table_name: str
    columns: tuple[str, ...]
    row_count: int
    header_row: int
    has_totals: bool


@dataclass(frozen=True)
class MergePreview:
    inputs: tuple[MergeInput, ...]
    target: Path
    prior_signature: FileSignature | None
    row_count: int


def _same_file(left: Path, right: Path) -> bool:
    return str(left).casefold() == str(right).casefold() or (
        left.exists() and right.exists() and left.samefile(right)
    )


def _validate_paths(sources: tuple[Path, ...], target: Path) -> None:
    if len(sources) < 2:
        raise WorkbookValidationError("병합할 .xlsx 파일을 2개 이상 선택하세요.")
    for index, source in enumerate(sources):
        _validate_source_path(source)
        if any(_same_file(source, other) for other in sources[:index]):
            raise WorkbookValidationError("같은 원본 파일을 중복 선택할 수 없습니다.")
        if _same_file(source, target):
            raise WorkbookValidationError("결과 파일은 원본 파일을 덮어쓸 수 없습니다.")
    _validate_output_dir(target.parent)
    if (
        target.suffix.casefold() != ".xlsx"
        or target.is_reserved()
        or _INVALID_CHARACTERS.search(target.name)
        or target.name.rstrip(" .") != target.name
    ):
        raise WorkbookValidationError("결과 파일에 유효한 .xlsx 파일명을 지정하세요.")
    if target.exists() and not target.is_file():
        raise WorkbookValidationError("결과 경로가 파일이 아닙니다.")
    if max(len(str(target)), len(str(target.parent / ".em-00000000.xlsx"))) > 218:
        raise WorkbookValidationError("결과 또는 Excel 임시 경로는 218자를 넘을 수 없습니다.")


def _validate_unchanged(preview: MergePreview) -> None:
    _validate_paths(tuple(item.source for item in preview.inputs), preview.target)
    for item in preview.inputs:
        if not same_signature(item.source, item.signature):
            raise WorkbookValidationError(f"원본 파일이 미리보기 이후 변경되었습니다: {item.source.name}")
    if not same_signature(preview.target, preview.prior_signature):
        raise WorkbookValidationError("결과 파일이 미리보기 이후 변경되었습니다.")


class MergeService:
    def preview(self, sources: tuple[Path, ...], target: Path) -> MergePreview:
        sources = tuple(Path(source).resolve() for source in sources)
        target = Path(target).resolve()
        _validate_paths(sources, target)
        prior = capture_signature(target) if target.exists() else None
        inputs: list[MergeInput] = []
        with _excel_session() as excel:
            # Inspect the template last so expansion needs no reopen or concurrent books.
            for source in sources[1:] + sources[:1]:
                signature = capture_signature(source)
                workbook = _open_workbook(excel, source, read_only=True)
                try:
                    item = _inspect_source(workbook, source, signature)
                    if source != sources[0]:
                        inputs.append(item)
                        continue
                    inputs.insert(0, item)
                    for other in inputs[1:]:
                        if other.columns != item.columns:
                            raise WorkbookValidationError(f"Table 열 이름과 순서가 다릅니다: {other.source.name}")
                    total = sum(item.row_count for item in inputs)
                    if item.header_row + total + int(item.has_totals) > 1_048_576:
                        raise WorkbookValidationError("병합 결과가 Excel 워크시트 행 한도를 초과합니다.")
                    if total > item.row_count:
                        sheet, table = _merge_table(workbook)
                        _validate_merge_expansion(sheet, table, total)
                finally:
                    _close_without_saving(workbook)
        preview = MergePreview(tuple(inputs), target, prior, total)
        _validate_unchanged(preview)
        return preview

    def execute(self, preview: MergePreview, overwrite: bool, progress: ProgressCallback) -> Path:
        progress(0, 0, "원본 확인 중")
        _validate_unchanged(preview)
        if preview.prior_signature is not None and not overwrite:
            raise WorkbookValidationError("기존 파일 덮어쓰기 승인이 필요합니다.")
        descriptor, filename = tempfile.mkstemp(prefix=".em-", suffix=".xlsx", dir=preview.target.parent)
        os.close(descriptor)
        temp = Path(filename)
        try:
            first = preview.inputs[0]
            progress(0, 0, "파일 복사 중")
            shutil.copy2(first.source, temp)
            temp.chmod(stat.S_IREAD | stat.S_IWRITE)
            if capture_signature(temp) != first.signature:
                raise WorkbookValidationError("첫 번째 원본 복사본이 미리보기와 다릅니다.")
            progress(0, 0, "값 병합 중")
            _write_merged(temp, preview, progress)
            _validate_unchanged(preview)
            _publish_temp(temp, preview.target, preview.prior_signature)
            progress(1, 1, "완료")
            return preview.target
        finally:
            active_error = sys.exception()
            try:
                temp.unlink(missing_ok=True)
            except OSError as exc:
                raise SplitExecutionError(f"병합 임시 파일을 삭제하지 못했습니다: {temp}; {active_error or exc}") from exc


def _table_headers(table: Any) -> tuple[str, ...]:
    values = table.HeaderRowRange.Value2
    return tuple(str(value) for value in (values[0] if isinstance(values, tuple) else (values,)))


def _inspect_source(workbook: Any, source: Path, signature: FileSignature) -> MergeInput:
    sheet, table = _merge_table(workbook)
    return MergeInput(
        source, signature, str(sheet.Name), str(table.Name), _table_headers(table),
        int(table.ListRows.Count), int(table.HeaderRowRange.Row), bool(table.ShowTotals),
    )


def _merge_table(workbook: Any) -> tuple[Any, Any]:
    if workbook.Worksheets.Count != 1 or workbook.Sheets.Count != 1:
        raise WorkbookValidationError("병합 원본마다 워크시트가 정확히 하나 있어야 합니다.")
    sheet = workbook.Worksheets.Item(1)
    if workbook.ProtectStructure or sheet.ProtectContents or sheet.Visible != -1:
        raise WorkbookValidationError("보호되거나 숨겨진 워크시트는 병합할 수 없습니다.")
    table = _single_table(sheet.ListObjects)
    if table.SourceType != 1:
        raise WorkbookValidationError("외부 연결 Table은 병합할 수 없습니다.")
    if not table.ShowHeaders:
        raise WorkbookValidationError("Table 머리글 행이 표시되어 있어야 합니다.")
    return sheet, table


def _validate_merge_expansion(sheet: Any, table: Any, row_count: int) -> None:
    first_row = int(table.Range.Row) + int(table.Range.Rows.Count)
    last_row = int(table.HeaderRowRange.Row) + row_count + int(table.ShowTotals)
    if first_row > last_row:
        return
    first_column = int(table.Range.Column)
    last_column = first_column + int(table.Range.Columns.Count) - 1
    bounded = sheet.Range(sheet.Cells(first_row, first_column), sheet.Cells(last_row, last_column))
    message = "병합 Table 확장 범위에 덮어쓸 수 없는 콘텐츠가 있습니다."
    if (
        _bulk_contains_value(bounded.Value2)
        or _bulk_contains_value(bounded.Formula)
        or bounded.MergeCells != False  # None denotes a mix of merged and unmerged cells.
        or int(bounded.Hyperlinks.Count) > 0
    ):
        raise WorkbookValidationError(message)
    # Formatting alone is safe; inspect only actual comment anchors, not every cell.
    for name in ("Comments", "CommentsThreaded"):
        try:
            comments = getattr(sheet, name)
            count = int(comments.Count)
        except Exception as exc:
            if name == "CommentsThreaded" and unsupported_threaded_comments_error(exc):
                continue
            raise
        for index in range(1, count + 1):
            anchor = comments.Item(index).Parent
            if first_row <= int(anchor.Row) <= last_row and first_column <= int(anchor.Column) <= last_column:
                raise WorkbookValidationError(message)


def _write_merged(temp: Path, preview: MergePreview, progress: ProgressCallback) -> None:
    with _excel_session() as excel:
        workbook = _open_workbook(excel, temp, read_only=False)
        try:
            sheet, table = _merge_table(workbook)
            if sheet.FilterMode:
                sheet.ShowAllData()
            columns = len(preview.inputs[0].columns)
            header_row = int(table.HeaderRowRange.Row)
            first_column = int(table.Range.Column)
            last_column = first_column + columns - 1
            if int(table.ListRows.Count) != preview.row_count:
                _validate_merge_expansion(sheet, table, preview.row_count)
                # Range.Resize is an optional-argument COM property; use explicit cells.
                table.Resize(sheet.Range(
                    sheet.Cells(header_row, first_column),
                    sheet.Cells(header_row + preview.row_count + int(table.ShowTotals), last_column),
                ))
            offset = 0
            for completed, item in enumerate(preview.inputs, 1):
                source = _open_workbook(excel, item.source, read_only=True)
                try:
                    source_sheet, source_table = _merge_table(source)
                    headers = _table_headers(source_table)
                    if headers != item.columns or int(source_table.ListRows.Count) != item.row_count:
                        raise WorkbookValidationError(f"원본 Table이 미리보기와 다릅니다: {item.source.name}")
                    if source_sheet.FilterMode:
                        source_sheet.ShowAllData()
                    if item.row_count:
                        destination = sheet.Range(
                            sheet.Cells(header_row + offset + 1, first_column),
                            sheet.Cells(header_row + offset + item.row_count, last_column),
                        )
                        source_table.DataBodyRange.Copy()
                        destination.PasteSpecial(Paste=-4122)  # xlPasteFormats
                        destination.PasteSpecial(Paste=12)  # xlPasteValuesAndNumberFormats
                        excel.CutCopyMode = False
                    offset += item.row_count
                finally:
                    _close_without_saving(source)
                progress(completed, len(preview.inputs), item.source.name)
            if int(table.ListRows.Count) != preview.row_count:
                raise SplitExecutionError("병합 결과 Table 행 수가 예상과 다릅니다.")
            if preview.row_count:
                # Freeze any calculated-column autofill while preserving error
                # cells and literal strings beginning with '=' as Excel values.
                table.DataBodyRange.Copy()
                table.DataBodyRange.PasteSpecial(Paste=-4163)  # xlPasteValues
                excel.CutCopyMode = False
            sheet.Calculate()
            progress(0, 0, "결과 저장 중")
            workbook.SaveAs(
                str(temp), FileFormat=51, Password="", WriteResPassword="",
                ReadOnlyRecommended=False, AddToMru=False,
            )
            if Path(str(workbook.FullName)).resolve() != temp.resolve():
                raise SplitExecutionError("병합 결과가 지정한 임시 경로에 저장되지 않았습니다.")
        finally:
            _close_without_saving(workbook)
        progress(0, 0, "저장 결과 확인 중")
        _verify_xlsx_package(temp)
