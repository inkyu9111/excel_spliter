from __future__ import annotations

import os
import logging
import shutil
import stat
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .errors import SplitExecutionError, WorkbookValidationError
from .excel_gateway import _close_without_saving, _com_stage, _excel_error_code, _excel_session, _open_workbook, _publish_temp
from .file_signature import capture_signature, same_signature
from .merge_service import _same_file
from .naming import _INVALID_CHARACTERS
from .ports import ProgressCallback
from .source_session import _verify_xlsx_package
from .split_service import _validate_output_dir, _validate_source_path


logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class CompareTable:
    sheet_name: str
    table_name: str
    columns: tuple[str, ...]


@dataclass(frozen=True)
class CompareDifference:
    kind: str
    sheet_name: str
    cell: str = ""
    column_name: str = ""
    key: str = ""
    reference_value: Any = None
    comparison_value: Any = None
    reference_cell: str = ""
    comparison_cell: str = ""


@dataclass(frozen=True)
class CompareResult:
    target: Path
    changed_cells: int
    missing_sheets: tuple[str, ...]
    missing_rows: int = 0
    missing_columns: tuple[str, ...] = ()
    added_rows: int = 0
    modified_cells: int = 0
    details: tuple[CompareDifference, ...] = ()
    details_truncated: bool = False
    omitted_details: int = 0


@dataclass
class _ComparisonDetails:
    items: list[CompareDifference] = field(default_factory=list)
    added_rows: int = 0
    modified_cells: int = 0
    omitted: int = 0

    def add(self, **values: Any) -> None:
        # ponytail: retain 1000 detail records; stream a separate export if all details are needed.
        if len(self.items) < 1000:
            self.items.append(CompareDifference(**values))
        else:
            self.omitted += 1


def _cell_address(row: int, column: int) -> str:
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def _key_label(keys: tuple[str, ...], key: tuple) -> str:
    return ", ".join(f"{name}={value!r}" for name, (_type, value) in zip(keys, key, strict=True))


def _validate_paths(reference: Path, comparison: Path, target: Path) -> None:
    for source in (reference, comparison):
        _validate_source_path(source)
        if _same_file(source, target):
            raise WorkbookValidationError("결과는 원본과 다른 새 파일에 저장해야 합니다.")
    _validate_output_dir(target.parent)
    if (
        target.suffix.casefold() != ".xlsx"
        or target.is_reserved()
        or _INVALID_CHARACTERS.search(target.name)
        or target.name.rstrip(" .") != target.name
    ):
        raise WorkbookValidationError("결과에 유효한 .xlsx 파일명을 지정하세요.")
    if target.exists() or target.is_symlink():
        raise WorkbookValidationError("결과 파일이 이미 존재합니다. 새 파일명을 지정하세요.")
    if max(len(str(target)), len(str(target.parent / ".ec-00000000.xlsx"))) > 218:
        raise WorkbookValidationError("결과 또는 Excel 임시 경로는 218자를 넘을 수 없습니다.")


class CompareService:
    def inspect_tables(self, reference: Path, comparison: Path) -> tuple[tuple[CompareTable, ...], tuple[CompareTable, ...]]:
        sources = (Path(reference).resolve(), Path(comparison).resolve())
        for source in sources:
            _validate_source_path(source)
        signatures = tuple(capture_signature(source) for source in sources)
        result = []
        with _excel_session() as excel:
            for source in sources:
                workbook = _open_workbook(excel, source, read_only=True)
                try:
                    tables = tuple(CompareTable(str(sheet.Name), str(table.Name), _columns(table)) for sheet, table in _tables(workbook))
                    if not tables:
                        raise WorkbookValidationError(f"Excel Table이 없습니다: {source.name}")
                    result.append(tables)
                finally:
                    _close_without_saving(workbook)
        for source, signature in zip(sources, signatures, strict=True):
            if not same_signature(source, signature):
                raise WorkbookValidationError(f"Table을 읽는 동안 원본이 변경되었습니다: {source.name}")
        return result[0], result[1]

    def execute(
        self, reference: Path, comparison: Path, target: Path, *, progress: ProgressCallback,
        key_columns: tuple[str, ...] = (),
        reference_table: tuple[str, str] | None = None,
        comparison_table: tuple[str, str] | None = None,
    ) -> CompareResult:
        if not key_columns and (reference_table is not None or comparison_table is not None):
            raise WorkbookValidationError("Table 비교에 사용할 Key 열을 하나 이상 선택하세요.")
        reference, comparison = Path(reference).resolve(), Path(comparison).resolve()
        target = Path(target).absolute()
        # Reject an existing symlink itself before resolving its destination.
        _validate_paths(reference, comparison, target)
        target = target.resolve()
        logger.info(
            "Compare start: mode=%s reference=%s comparison=%s target=%s",
            "key" if key_columns else "position", reference, comparison, target,
        )
        signatures = tuple(capture_signature(source) for source in (reference, comparison))
        descriptor, filename = tempfile.mkstemp(prefix=".ec-", suffix=".xlsx", dir=target.parent)
        os.close(descriptor)
        temp = Path(filename)
        try:
            progress(0, 0, "비교 파일 복사")
            shutil.copy2(comparison, temp)
            temp.chmod(stat.S_IREAD | stat.S_IWRITE)
            if capture_signature(temp) != signatures[1]:
                raise WorkbookValidationError("비교 파일을 복사하는 동안 변경되었습니다.")
            missing_rows, missing_columns = 0, ()
            details = _ComparisonDetails()
            if key_columns:
                changed, missing_rows, missing_columns = _write_key_comparison(
                    reference, comparison, temp, progress, key_columns, reference_table, comparison_table, details,
                )
                missing = ()
            else:
                changed, missing = _write_comparison(reference, temp, progress, details)
            progress(0, 0, "결과 검증")
            _verify_xlsx_package(temp)
            _validate_paths(reference, comparison, target)
            for source, signature in zip((reference, comparison), signatures, strict=True):
                if not same_signature(source, signature):
                    raise WorkbookValidationError(f"비교하는 동안 원본이 변경되었습니다: {source.name}")
            _publish_temp(temp, target, None)
            progress(1, 1, "완료")
            return CompareResult(target, changed, missing, missing_rows, missing_columns,
                details.added_rows, details.modified_cells, tuple(details.items), bool(details.omitted), details.omitted)
        finally:
            active_error = sys.exception()
            try:
                temp.unlink(missing_ok=True)
            except OSError as exc:
                raise SplitExecutionError(f"비교 임시 파일을 삭제하지 못했습니다: {temp}; {active_error or exc}") from exc


def _range_rows(sheet: Any, used: Any, *, sheet_name: str = "(이름 없음)"):
    first_row, first_column = int(used.Row), int(used.Column)
    row_count, column_count = int(used.Rows.Count), int(used.Columns.Count)
    last_column = first_column + column_count - 1
    rows_per_chunk = max(1, 32768 // column_count)
    for row in range(first_row, first_row + row_count, rows_per_chunk):
        last_row = min(row + rows_per_chunk, first_row + row_count) - 1
        with _com_stage(f"값 읽기: {sheet_name} R{row}C{first_column}:R{last_row}C{last_column}"):
            block = sheet.Range(sheet.Cells(row, first_column), sheet.Cells(last_row, last_column)).Value2
        if last_row == row and column_count == 1:
            block = ((block,),)
        for row_offset, cells in enumerate(block):
            yield row + row_offset, cells


def _read_values(sheet: Any) -> dict[tuple[int, int], Any]:
    used = sheet.UsedRange
    first_column = int(used.Column)
    first_row, row_count, column_count = int(used.Row), int(used.Rows.Count), int(used.Columns.Count)
    sheet_name = str(getattr(sheet, "Name", "(이름 없음)"))
    logger.info(
        "Compare read: sheet=%s used_range=R%sC%s rows=%s columns=%s",
        sheet_name, first_row, first_column, row_count, column_count,
    )
    # ponytail: keep snapshots in memory; spill them if workbook size requires it.
    return {
        (row, first_column + offset): value
        for row, cells in _range_rows(sheet, used, sheet_name=sheet_name)
        for offset, value in enumerate(cells)
        if value is not None and value != ""
    }


def _typed_value(value: Any) -> tuple[Any, Any]:
    if value is None or value == "":
        return (None, None)
    error = _excel_error_code(value)
    if error is not None:
        return ("error", error)
    if type(value) in (int, float):
        return (float, value)
    return (type(value), value)


def _columns(table: Any) -> tuple[str, ...]:
    return tuple(str(table.ListColumns.Item(index).Name) for index in range(1, table.ListColumns.Count + 1))


def _tables(workbook: Any):
    for index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets.Item(index)
        for table_index in range(1, sheet.ListObjects.Count + 1):
            yield sheet, sheet.ListObjects.Item(table_index)


def _select_table(workbook: Any, selector: tuple[str, str] | None, source: Path) -> tuple[Any, Any]:
    tables = list(_tables(workbook))
    if selector is None:
        if len(tables) != 1:
            raise WorkbookValidationError(f"비교할 Table을 하나 선택하세요: {source.name} (Table {len(tables)}개)")
        return tables[0]
    for sheet, table in tables:
        if (str(sheet.Name), str(table.Name)) == selector:
            return sheet, table
    raise WorkbookValidationError(f"선택한 Table이 없습니다: {source.name} / {selector[0]} / {selector[1]}")


def _keyed_rows(sheet: Any, table: Any, keys: tuple[str, ...], source: Path):
    columns = _columns(table)
    absent = tuple(key for key in keys if key not in columns)
    if absent:
        raise WorkbookValidationError(f"Key 열이 없습니다: {source.name} / {table.Name}: {', '.join(absent)}")
    indexes = tuple(columns.index(key) for key in keys)
    rows = {}
    if int(table.ListRows.Count):
        for row, cells in _range_rows(sheet, table.DataBodyRange, sheet_name=str(getattr(sheet, "Name", "(이름 없음)"))):
            key = tuple(_typed_value(cells[index]) for index in indexes)
            if key in rows:
                raise WorkbookValidationError(f"Key가 중복되었습니다: {source.name} / {sheet.Name} / {table.Name}, 행 {rows[key][0]}, {row}")
            rows[key] = row, cells
    return columns, rows


def _highlight(sheet: Any, coordinates: list[tuple[int, int]]) -> None:
    if not coordinates:
        return
    if sheet.ProtectContents:
        raise WorkbookValidationError(f"보호된 시트에 비교 표시를 할 수 없습니다: {sheet.Name}")
    start_row, start_column = coordinates[0]
    sheet_name = str(getattr(sheet, "Name", "(이름 없음)"))
    end_column = start_column
    for row, column in coordinates[1:] + [(0, 0)]:
        if row == start_row and column == end_column + 1:
            end_column = column
            continue
        range_name = f"R{start_row}C{start_column}:R{start_row}C{end_column}"
        with _com_stage(f"강조 직접 채우기: {sheet_name} {range_name}"):
            cells = sheet.Range(sheet.Cells(start_row, start_column), sheet.Cells(start_row, end_column))
            cells.Interior.Pattern = 1  # xlSolid
            cells.Interior.Color = 65535  # RGB(255, 255, 0)
        with _com_stage(f"강조 조건부 서식 확인: {sheet_name} {range_name}"):
            has_conditions = int(cells.FormatConditions.Count)
        if has_conditions:
            with _com_stage(f"강조 조건부 서식 추가: {sheet_name} {range_name}"):
                # Pass all positional slots to avoid late-bound optional-argument errors.
                rule = cells.FormatConditions.Add(2, 3, "=TRUE", "")
                rule.Interior.Color = 65535
            with _com_stage(f"강조 조건부 서식 우선순위: {sheet_name} {range_name}"):
                rule.SetFirstPriority()
                rule.StopIfTrue = False  # Keep lower-priority fonts and borders.
        start_row, start_column, end_column = row, column, column


def _write_comparison(
    reference: Path, temp: Path, progress: ProgressCallback, details: _ComparisonDetails | None = None,
) -> tuple[int, tuple[str, ...]]:
    details = details if details is not None else _ComparisonDetails()
    progress(0, 0, "기준 파일 열기")
    with _excel_session() as excel:
        with _com_stage(f"기준 파일 열기: {reference.name}"):
            baseline = _open_workbook(excel, reference, read_only=True)
        try:
            with _com_stage(f"기준 파일 계산: {reference.name}"):
                excel.Calculate()
            with _com_stage(f"기준 시트 검색: {reference.name}"):
                sheets = tuple(baseline.Worksheets.Item(index) for index in range(1, baseline.Worksheets.Count + 1))
            values = {}
            for index, sheet in enumerate(sheets):
                progress(index, len(sheets), f"기준 값 읽기: {sheet.Name}")
                values[str(sheet.Name)] = _read_values(sheet)
            progress(len(sheets), len(sheets), "기준 값 읽기")
        finally:
            _close_without_saving(baseline)
        progress(0, 0, "비교 파일 열기")
        with _com_stage(f"비교 파일 열기: {temp.name}"):
            workbook = _open_workbook(excel, temp, read_only=False)
        try:
            with _com_stage(f"비교 파일 계산: {temp.name}"):
                excel.Calculate()
            changed = 0
            with _com_stage(f"비교 시트 검색: {temp.name}"):
                total = int(workbook.Worksheets.Count)
            for index in range(1, total + 1):
                with _com_stage(f"비교 시트 검색: {temp.name} #{index}"):
                    sheet = workbook.Worksheets.Item(index)
                    sheet_name = str(sheet.Name)
                is_added_sheet = sheet_name not in values
                previous = values.pop(sheet_name, {})
                progress(0, 0, f"비교 값 읽기: {sheet_name}")
                current = _read_values(sheet)
                progress(0, 0, f"차이 비교: {sheet_name}")
                differences = sorted(
                    coordinate for coordinate in previous.keys() | current.keys()
                    if _typed_value(previous.get(coordinate)) != _typed_value(current.get(coordinate))
                )
                for row, column in differences:
                    coordinate = (row, column)
                    kind = "added" if coordinate not in previous else "missing" if coordinate not in current else "changed"
                    details.modified_cells += kind == "changed"
                    cell = _cell_address(row, column)
                    details.add(kind=kind, sheet_name=sheet_name, cell=cell,
                        reference_value=previous.get(coordinate), comparison_value=current.get(coordinate),
                        reference_cell=cell if coordinate in previous else "",
                        comparison_cell=cell if coordinate in current else "")
                if is_added_sheet and not current:
                    details.add(kind="added", sheet_name=sheet_name, comparison_value="시트 추가")
                logger.info("Compare sheet: sheet=%s differences=%s", sheet_name, len(differences))
                progress(0, 0, f"노란색 표시: {sheet_name}")
                _highlight(sheet, differences)
                changed += len(differences)
                progress(index, total, f"시트 비교: {sheet_name}")
            for sheet_name in values:
                details.add(kind="missing", sheet_name=sheet_name, reference_value="시트 누락")
            progress(0, 0, "결과 저장")
            _save_comparison(workbook, temp)
        finally:
            _close_without_saving(workbook)
    return changed, tuple(values)


def _save_comparison(workbook: Any, temp: Path) -> None:
    with _com_stage(f"결과 저장: {temp.name}"):
        workbook.SaveAs(
            str(temp), FileFormat=51, Password="", WriteResPassword="",
            ReadOnlyRecommended=False, AddToMru=False,
        )
    if Path(str(workbook.FullName)).resolve() != temp:
        raise SplitExecutionError("결과가 지정한 임시 경로에 저장되지 않았습니다.")


def _write_key_comparison(
    reference: Path, comparison: Path, temp: Path, progress: ProgressCallback, keys: tuple[str, ...],
    reference_table: tuple[str, str] | None, comparison_table: tuple[str, str] | None,
    details: _ComparisonDetails | None = None,
) -> tuple[int, int, tuple[str, ...]]:
    details = details if details is not None else _ComparisonDetails()
    progress(0, 0, "기준 파일 열기")
    with _excel_session() as excel:
        with _com_stage(f"기준 파일 열기: {reference.name}"):
            baseline = _open_workbook(excel, reference, read_only=True)
        try:
            with _com_stage(f"기준 파일 계산: {reference.name}"):
                excel.Calculate()
            with _com_stage(f"기준 Table 검색: {reference.name}"):
                sheet, table = _select_table(baseline, reference_table, reference)
            progress(0, 0, f"기준 값 읽기: {sheet.Name} / {table.Name}")
            previous_columns, previous = _keyed_rows(sheet, table, keys, reference)
            previous_sheet_name = str(sheet.Name)
            previous_first_column = int(table.DataBodyRange.Column) if previous else 0
            header = table.HeaderRowRange
            previous_header = (int(header.Row), int(header.Column)) if header is not None else None
        finally:
            _close_without_saving(baseline)
        progress(0, 0, "비교 파일 열기")
        with _com_stage(f"비교 파일 열기: {temp.name}"):
            workbook = _open_workbook(excel, temp, read_only=False)
        try:
            with _com_stage(f"비교 파일 계산: {temp.name}"):
                excel.Calculate()
            with _com_stage(f"비교 Table 검색: {comparison.name}"):
                sheet, table = _select_table(workbook, comparison_table, comparison)
            progress(0, 0, f"비교 값 읽기: {sheet.Name} / {table.Name}")
            columns, current = _keyed_rows(sheet, table, keys, comparison)
            first_column = int(table.DataBodyRange.Column) if current else 0
            previous_indexes = {name: index for index, name in enumerate(previous_columns)}
            differences = []
            progress(0, len(current), "차이 비교")
            for key, (row, cells) in current.items():
                prior = previous.get(key)
                label = _key_label(keys, key)
                if prior is None:
                    details.added_rows += 1
                    cell = _cell_address(row, first_column)
                    details.add(kind="added", sheet_name=str(sheet.Name), cell=cell, key=label,
                        comparison_value=tuple(cells), comparison_cell=cell)
                for index, name in enumerate(columns):
                    prior_value = prior[1][previous_indexes[name]] if prior is not None and name in previous_indexes else None
                    if prior is None or _typed_value(prior_value) != _typed_value(cells[index]):
                        differences.append((row, first_column + index))
                        if prior is not None:
                            kind = "changed" if name in previous_indexes else "added"
                            details.modified_cells += kind == "changed"
                            cell = _cell_address(row, first_column + index)
                            details.add(kind=kind, sheet_name=str(sheet.Name), cell=cell, column_name=name, key=label,
                                reference_value=prior_value, comparison_value=cells[index], comparison_cell=cell,
                                reference_cell=_cell_address(prior[0], previous_first_column + previous_indexes[name])
                                    if name in previous_indexes else "")
            progress(len(current), len(current), "차이 비교")
            for key, (row, cells) in previous.items():
                if key not in current:
                    cell = _cell_address(row, previous_first_column)
                    details.add(kind="missing", sheet_name=previous_sheet_name, cell=cell, key=_key_label(keys, key),
                        reference_value=tuple(cells), reference_cell=cell)
            for name in previous_columns:
                if name not in columns:
                    cell = _cell_address(previous_header[0], previous_header[1] + previous_indexes[name]) if previous_header else ""
                    details.add(kind="missing", sheet_name=previous_sheet_name, column_name=name, cell=cell,
                        reference_value="열 누락", reference_cell=cell)
            # Both Tables have passed key validation before any cell is formatted.
            logger.info("Compare key sheet: sheet=%s differences=%s", sheet.Name, len(differences))
            progress(0, 0, f"노란색 표시: {sheet.Name} / {table.Name}")
            _highlight(sheet, differences)
            progress(0, 0, "결과 저장")
            _save_comparison(workbook, temp)
        finally:
            _close_without_saving(workbook)
    return len(differences), len(previous.keys() - current.keys()), tuple(name for name in previous_columns if name not in columns)
