import pytest

from excel_splitter.classifier import canonicalize, group_samples
from excel_splitter.errors import WorkbookValidationError
from excel_splitter.models import CanonicalKey, CellSample


def test_blank_and_formula_empty_share_one_group() -> None:
    groups = group_samples([CellSample(1, None, ""), CellSample(2, "", "")])

    assert groups[0].key == CanonicalKey("blank", "")
    assert groups[0].row_indexes == (1, 2)


def test_bool_is_not_number_but_one_equals_one_point_zero() -> None:
    assert canonicalize(CellSample(1, True, "TRUE")) != canonicalize(
        CellSample(2, 1, "1")
    )
    assert canonicalize(CellSample(1, 1, "1")) == canonicalize(
        CellSample(2, 1.0, "1.0")
    )


def test_error_code_is_part_of_key() -> None:
    assert canonicalize(CellSample(1, None, "#N/A", True, 2042)) == CanonicalKey(
        "error", "2042"
    )


def test_error_without_code_is_rejected() -> None:
    with pytest.raises(WorkbookValidationError, match="Excel 오류 코드"):
        canonicalize(CellSample(1, None, "#N/A", True))


def test_text_preserves_case_and_whitespace() -> None:
    assert canonicalize(CellSample(1, " Seoul ", " Seoul ")) != canonicalize(
        CellSample(2, "seoul", "seoul")
    )


def test_unsupported_value_type_is_rejected() -> None:
    with pytest.raises(WorkbookValidationError, match="지원하지 않는 Excel 값 형식"):
        canonicalize(CellSample(1, object(), "object"))


def test_groups_keep_first_seen_order_and_first_display_text() -> None:
    groups = group_samples(
        [
            CellSample(7, 2, "둘"),
            CellSample(9, 1, "하나"),
            CellSample(11, 2.0, "2.0"),
        ]
    )

    assert [(group.label, group.count, group.row_indexes) for group in groups] == [
        ("둘", 2, (7, 11)),
        ("하나", 1, (9,)),
    ]


def test_unusable_display_text_falls_back_by_value_type() -> None:
    groups = group_samples(
        [
            CellSample(1, "원문", "####"),
            CellSample(2, True, ""),
            CellSample(3, 1.25, "####"),
        ]
    )

    assert [group.label for group in groups] == ["원문", "TRUE", "1.25"]
