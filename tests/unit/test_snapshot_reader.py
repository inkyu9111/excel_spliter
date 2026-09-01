from __future__ import annotations

import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.models import CanonicalKey, CellSample
from excel_splitter.snapshot_reader import (
    build_samples,
    group_bulk_samples,
    normalize_column_values,
)


class FakeCell:
    def __init__(self, text: str) -> None:
        self.Text = text


class FakeCells:
    def __init__(self, texts: tuple[str, ...]) -> None:
        self._texts = texts
        self.calls: list[tuple[int, int]] = []

    def Item(self, row: int, column: int) -> FakeCell:
        self.calls.append((row, column))
        return FakeCell(self._texts[row - 1])


class FakeRange:
    def __init__(self, value2: object, texts: tuple[str, ...]) -> None:
        self._value2 = value2
        self.value2_reads = 0
        self.Cells = FakeCells(texts)

    @property
    def Value2(self) -> object:
        self.value2_reads += 1
        return self._value2


def test_normalize_column_values_maps_one_row_scalar() -> None:
    assert normalize_column_values("서울", 1) == ("서울",)


def test_normalize_column_values_maps_multi_row_one_column_matrix() -> None:
    assert normalize_column_values((("서울",), ("부산",), ("대구",)), 3) == (
        "서울",
        "부산",
        "대구",
    )


@pytest.mark.parametrize(
    ("value2", "row_count"),
    [
        ((1, 2), 2),
        (((1, 2),), 1),
        (((1,),), 2),
        ("서울", 2),
        (((1,),), 0),
    ],
)
def test_normalize_column_values_rejects_wrong_shape(
    value2: object, row_count: int
) -> None:
    with pytest.raises(WorkbookValidationError):
        normalize_column_values(value2, row_count)


def test_build_samples_reads_value_once_and_emits_one_based_error_samples() -> None:
    data_range = FakeRange(
        ((None,), (-2146826246,), (True,), (1.25,), ("서울",)),
        ("", "#N/A", "TRUE", "1.25", "서울"),
    )

    samples = build_samples(data_range, 5)

    assert data_range.value2_reads == 1
    assert samples == [
        CellSample(1, None, ""),
        CellSample(2, -2146826246, "#N/A", True, 2042),
        CellSample(3, True, "TRUE"),
        CellSample(4, 1.25, "1.25"),
        CellSample(5, "서울", "서울"),
    ]


def test_group_bulk_samples_keeps_order_and_reads_text_once_per_unique_key() -> None:
    data_range = FakeRange(
        ((2,), (1,), (2.0,), (None,), ("",), (True,), (-2146826246,)),
        ("둘", "하나", "2.0", "", "", "TRUE", "#N/A"),
    )

    groups = group_bulk_samples(data_range, 7)

    assert [group.key for group in groups] == [
        CanonicalKey("number", "2"),
        CanonicalKey("number", "1"),
        CanonicalKey("blank", ""),
        CanonicalKey("bool", "true"),
        CanonicalKey("error", "2042"),
    ]
    assert [group.label for group in groups] == ["둘", "하나", "", "TRUE", "#N/A"]
    assert [group.row_indexes for group in groups] == [
        (1, 3),
        (2,),
        (4, 5),
        (6,),
        (7,),
    ]
    assert data_range.value2_reads == 1
    assert data_range.Cells.calls == [(1, 1), (2, 1), (4, 1), (6, 1), (7, 1)]
