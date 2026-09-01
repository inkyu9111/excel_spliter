from collections import OrderedDict
from decimal import Decimal

from .errors import WorkbookValidationError
from .models import CanonicalKey, CellSample, GroupSummary


_ERROR_TEXT = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
    2043: "#GETTING_DATA",
}


def canonicalize(sample: CellSample) -> CanonicalKey:
    if sample.is_error:
        if sample.error_code is None:
            raise WorkbookValidationError("Excel 오류 코드를 읽을 수 없습니다.")
        return CanonicalKey("error", str(sample.error_code))
    if sample.value is None or sample.value == "":
        return CanonicalKey("blank", "")
    if isinstance(sample.value, bool):
        return CanonicalKey("bool", "true" if sample.value else "false")
    if isinstance(sample.value, (int, float)):
        number = Decimal(str(sample.value)).normalize()
        if number == 0:
            number = Decimal(0)
        return CanonicalKey("number", format(number, "f"))
    if isinstance(sample.value, str):
        return CanonicalKey("text", sample.value)
    raise WorkbookValidationError(
        f"지원하지 않는 Excel 값 형식입니다: {type(sample.value).__name__}"
    )


def group_samples(samples: list[CellSample]) -> tuple[GroupSummary, ...]:
    grouped: OrderedDict[CanonicalKey, list[CellSample]] = OrderedDict()
    for sample in samples:
        grouped.setdefault(canonicalize(sample), []).append(sample)
    return tuple(
        GroupSummary(
            key=key,
            label=_label(rows[0], key),
            count=len(rows),
            row_indexes=tuple(row.row_index for row in rows),
        )
        for key, rows in grouped.items()
    )


def _label(sample: CellSample, key: CanonicalKey) -> str:
    if key.kind == "blank":
        return ""
    if sample.text and sample.text.strip().strip("#"):
        return sample.text
    if isinstance(sample.value, str):
        return sample.value
    if sample.is_error and sample.error_code is not None:
        return _ERROR_TEXT.get(sample.error_code, str(sample.error_code))
    if isinstance(sample.value, bool):
        return "TRUE" if sample.value else "FALSE"
    if isinstance(sample.value, (int, float)):
        return format(sample.value, ".15g")
    return key.value
