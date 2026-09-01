from __future__ import annotations

from typing import Any

from .classifier import canonicalize, group_samples
from .errors import WorkbookValidationError
from .models import CanonicalKey, CellSample, GroupSummary


_EXCEL_ERROR_CODES = {2000, 2007, 2015, 2023, 2029, 2036, 2042, 2043, 2045}


def normalize_column_values(value2: object, row_count: int) -> tuple[object, ...]:
    if row_count <= 0:
        raise WorkbookValidationError("Table 행 수가 올바르지 않습니다.")

    if not isinstance(value2, tuple):
        if row_count == 1:
            return (value2,)
        raise WorkbookValidationError("분류 컬럼의 행 수가 Table과 일치하지 않습니다.")

    if len(value2) != row_count:
        raise WorkbookValidationError("분류 컬럼의 행 수가 Table과 일치하지 않습니다.")
    if any(not isinstance(row, tuple) or len(row) != 1 for row in value2):
        raise WorkbookValidationError("분류 컬럼은 한 개 열이어야 합니다.")
    return tuple(row[0] for row in value2)


def build_samples(data_body_range: object, row_count: int) -> list[CellSample]:
    values = normalize_column_values(data_body_range.Value2, row_count)
    temporary = [_sample(index, value, "") for index, value in enumerate(values, 1)]

    representative_text: dict[CanonicalKey, str] = {}
    for sample in temporary:
        key = canonicalize(sample)
        if key not in representative_text:
            cell = data_body_range.Cells.Item(sample.row_index, 1)
            representative_text[key] = str(cell.Text)

    return [
        _sample(sample.row_index, sample.value, representative_text[canonicalize(sample)])
        for sample in temporary
    ]


def group_bulk_samples(
    data_body_range: object, row_count: int
) -> tuple[GroupSummary, ...]:
    return group_samples(build_samples(data_body_range, row_count))


def _sample(row_index: int, value: object, text: str) -> CellSample:
    error_code = _excel_error_code(value)
    return CellSample(
        row_index=row_index,
        value=value,
        text=text,
        is_error=error_code is not None,
        error_code=error_code,
    )


def _excel_error_code(value: Any) -> int | None:
    if type(value) is int:
        code = value & 0xFFFF if value < 0 else value
        if code in _EXCEL_ERROR_CODES:
            return code
    return None
