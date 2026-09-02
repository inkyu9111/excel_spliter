from contextlib import nullcontext
from types import SimpleNamespace

import pytest

from excel_splitter.errors import WorkbookValidationError


def _setup(tmp_path, monkeypatch):
    from excel_splitter import compare_service as compare

    reference = tmp_path / "reference.xlsx"
    comparison = tmp_path / "comparison.xlsx"
    target = tmp_path / "result.xlsx"
    reference.write_bytes(b"reference")
    comparison.write_bytes(b"comparison")
    monkeypatch.setattr(compare, "_verify_xlsx_package", lambda _: None)
    return compare, reference, comparison, target


def test_execute_publishes_only_the_modified_copy(tmp_path, monkeypatch):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)

    def write(source, temp, progress):
        assert source == reference
        assert temp != target and temp.read_bytes() == b"comparison"
        temp.write_bytes(b"highlighted comparison")
        progress(1, 1, "Data")
        return 3, ("Missing",)

    monkeypatch.setattr(compare, "_write_comparison", write)
    progress = []
    result = compare.CompareService().execute(reference, comparison, target, progress=lambda *args: progress.append(args))
    assert (result.target, result.changed_cells, result.missing_sheets) == (target, 3, ("Missing",))
    assert target.read_bytes() == b"highlighted comparison"
    assert reference.read_bytes() == b"reference" and comparison.read_bytes() == b"comparison"
    assert set(tmp_path.iterdir()) == {reference, comparison, target}
    assert progress == [(1, 1, "Data")]


@pytest.mark.parametrize("target_kind", ["reference", "comparison", "existing", "hardlink", "symlink", "dangling_symlink", "directory", "extension", "brackets"])
def test_rejects_unsafe_output_before_excel(tmp_path, monkeypatch, target_kind):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)
    if target_kind == "reference":
        target = reference
    elif target_kind == "comparison":
        target = comparison
    elif target_kind == "existing":
        target.write_bytes(b"prior output")
    elif target_kind == "hardlink":
        target.hardlink_to(comparison)
    elif target_kind == "symlink":
        target.symlink_to(comparison)
    elif target_kind == "dangling_symlink":
        target.symlink_to(tmp_path / "missing.xlsx")
    elif target_kind == "directory":
        target.mkdir()
    elif target_kind == "extension":
        target = target.with_suffix(".xlsm")
    else:
        target = tmp_path / "result[1].xlsx"
    monkeypatch.setattr(compare, "_write_comparison", lambda *_: pytest.fail("unsafe path reached Excel"))
    with pytest.raises(WorkbookValidationError):
        compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None)
    assert reference.read_bytes() == b"reference" and comparison.read_bytes() == b"comparison"


@pytest.mark.parametrize("failure", ["writer", "reference_changed", "comparison_changed", "target_created", "package"])
def test_failure_cleans_partial_output_without_overwriting_inputs(tmp_path, monkeypatch, failure):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)

    def write(_source, temp, _progress):
        temp.write_bytes(b"partial")
        if failure == "writer":
            raise RuntimeError("Excel failed")
        if failure == "reference_changed":
            reference.write_bytes(b"external change")
        elif failure == "comparison_changed":
            comparison.write_bytes(b"external change")
        elif failure == "target_created":
            target.write_bytes(b"new owner")
        return 1, ()

    monkeypatch.setattr(compare, "_write_comparison", write)
    if failure == "package":
        monkeypatch.setattr(compare, "_verify_xlsx_package", lambda _: (_ for _ in ()).throw(RuntimeError("invalid package")))
    with pytest.raises((RuntimeError, WorkbookValidationError, OSError)):
        compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None)
    assert reference.read_bytes() == (b"external change" if failure == "reference_changed" else b"reference")
    assert comparison.read_bytes() == (b"external change" if failure == "comparison_changed" else b"comparison")
    assert set(tmp_path.iterdir()) == ({reference, comparison, target} if failure == "target_created" else {reference, comparison})
    if target.exists():
        assert target.read_bytes() == b"new owner"


def test_same_source_is_allowed_with_a_distinct_output(tmp_path, monkeypatch):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)
    monkeypatch.setattr(compare, "_write_comparison", lambda *_: (0, ()))
    result = compare.CompareService().execute(reference, reference, target, progress=lambda *_: None)
    assert result.changed_cells == 0 and target.read_bytes() == b"reference"


def test_copy_changed_during_copy_is_rejected_before_excel(tmp_path, monkeypatch):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)
    monkeypatch.setattr(compare.shutil, "copy2", lambda _source, temp: temp.write_bytes(b"corrupted copy"))
    monkeypatch.setattr(compare, "_write_comparison", lambda *_: pytest.fail("corrupted copy reached Excel"))
    with pytest.raises(WorkbookValidationError):
        compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None)
    assert not target.exists()
    assert set(tmp_path.iterdir()) == {reference, comparison}


def test_writer_compares_types_coordinates_and_sheet_names_then_saves_copy(tmp_path, monkeypatch):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)
    source_sheet = SimpleNamespace(Name="Data", values={
        (1, 1): True, (1, 2): "1", (1, 3): "same", (1, 4): " x", (2, 1): 8,
        (8, 6): "removed", (9, 1): -2146826281,
    })
    current_sheet = SimpleNamespace(Name="Data", values={
        (1, 1): 1.0, (1, 2): 1.0, (1, 3): "same", (1, 4): "x", (2, 1): 8.0,
        (3, 1): "added", (9, 1): -2146826281.0,
    })
    absent = SimpleNamespace(Name="Missing", values={(1, 1): "missing"})
    added = SimpleNamespace(Name="Added", values={(2, 2): "added"})
    events = []

    def book(sheets, name):
        def save(path, **options):
            assert path == str(target) and options["FileFormat"] == 51
            assert options["Password"] == "" and options["WriteResPassword"] == ""
            events.append(f"save {name}")

        return SimpleNamespace(
            Worksheets=SimpleNamespace(Count=len(sheets), Item=lambda index: sheets[index - 1]),
            FullName=str(target), ReadOnly=False,
            Save=lambda: pytest.fail("SaveAs is required for a plain xlsx result"), SaveAs=save,
            Close=lambda **_: events.append(f"close {name}"),
        )

    baseline, output = book([source_sheet, absent], "baseline"), book([added, current_sheet], "copy")

    def open_book(_excel, path, *, read_only):
        if read_only:
            assert path == reference
            return baseline
        assert "close baseline" in events and path == target
        return output

    monkeypatch.setattr(compare, "_excel_session", lambda: nullcontext(SimpleNamespace(Calculate=lambda: None)))
    monkeypatch.setattr(compare, "_open_workbook", open_book)
    monkeypatch.setattr(compare, "_read_values", lambda sheet: sheet.values)
    highlights = {}
    monkeypatch.setattr(compare, "_highlight", lambda sheet, cells: highlights.update({sheet.Name: cells}))
    changed, missing = compare._write_comparison(reference, target, lambda *_: None)
    assert changed == 7 and missing == ("Missing",)
    assert highlights == {
        "Data": [(1, 1), (1, 2), (1, 4), (3, 1), (8, 6), (9, 1)],
        "Added": [(2, 2)],
    }
    assert events == ["close baseline", "save copy", "close copy"]


@pytest.mark.parametrize("scalar", [False, True])
def test_reader_keeps_offsets_and_nonblank_types_with_bounded_reads(monkeypatch, scalar):
    from excel_splitter import compare_service as compare

    ranges = []
    values = {(5, 3): True, (5, 4): "", (6, 3): " ", (6, 4): 0, (33005, 3): "last"}

    def read(first, last):
        ranges.append((first, last))
        if first == last:
            return SimpleNamespace(Value2=values.get(first))
        return SimpleNamespace(Value2=tuple(
            tuple(values.get((row, column)) for column in range(first[1], last[1] + 1))
            for row in range(first[0], last[0] + 1)
        ))

    class UsedRange:
        Row = 5
        Rows = SimpleNamespace(Count=1 if scalar else 33001)
        Columns = SimpleNamespace(Count=1 if scalar else 2)
        column_reads = 0

        @property
        def Column(self):
            self.column_reads += 1
            assert self.column_reads <= 2, "Column metadata must not cross COM once per populated cell"
            return 3

    sheet = SimpleNamespace(
        UsedRange=UsedRange(),
        Cells=lambda row, column: (row, column), Range=read,
    )
    assert compare._read_values(sheet) == ({(5, 3): True} if scalar else {(5, 3): True, (6, 3): " ", (6, 4): 0, (33005, 3): "last"})
    assert all((last[0] - first[0] + 1) * (last[1] - first[1] + 1) <= 32768 for first, last in ranges)


def test_highlight_uses_row_runs_and_overrides_existing_conditional_fill():
    from excel_splitter import compare_service as compare

    ranges = []
    rule = SimpleNamespace(Interior=SimpleNamespace(), SetFirstPriority=lambda: None)
    conditions = SimpleNamespace(Count=1, Add=lambda **_: rule)

    def bounded(first, last):
        result = SimpleNamespace(Interior=SimpleNamespace(), FormatConditions=conditions)
        ranges.append((first, last, result))
        return result

    sheet = SimpleNamespace(ProtectContents=False, Cells=lambda row, column: (row, column), Range=bounded)
    compare._highlight(sheet, [(1, 1), (1, 2), (1, 4), (2, 1)])
    assert [(first, last) for first, last, _ in ranges] == [((1, 1), (1, 2)), ((1, 4), (1, 4)), ((2, 1), (2, 1))]
    assert all(result.Interior.Color == 65535 and result.Interior.Pattern == 1 for _, _, result in ranges)
    assert rule.Interior.Color == 65535
    assert rule.StopIfTrue is False


def test_protected_changed_sheet_fails_before_any_highlight():
    from excel_splitter import compare_service as compare

    sheet = SimpleNamespace(ProtectContents=True, Name="Protected", Range=lambda *_: pytest.fail("protected sheet changed"))
    with pytest.raises(WorkbookValidationError):
        compare._highlight(sheet, [(1, 1)])
    compare._highlight(sheet, [])


class _Collection:
    def __init__(self, *items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, index):
        return self.items[index - 1]


def _key_books(tmp_path, monkeypatch, reference_rows=None, comparison_rows=None):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)
    events, highlights = [], {}

    def make_sheet(name, table_name, columns, rows, first_row, first_column):
        body = SimpleNamespace(Row=first_row, Column=first_column, Rows=SimpleNamespace(Count=len(rows)), Columns=SimpleNamespace(Count=len(columns)))
        table = SimpleNamespace(Name=table_name, ListColumns=_Collection(*(SimpleNamespace(Name=name) for name in columns)),
                                ListRows=SimpleNamespace(Count=len(rows)), DataBodyRange=body if rows else None)

        def read(first, last):
            selected = tuple(tuple(rows[row - first_row][column - first_column] for column in range(first[1], last[1] + 1))
                             for row in range(first[0], last[0] + 1))
            return SimpleNamespace(Value2=selected[0][0] if first == last else selected)

        return SimpleNamespace(Name=name, ListObjects=_Collection(table), Cells=lambda row, column: (row, column), Range=read)

    baseline_sheet = make_sheet("Baseline", "OldTable", ("Part", "Code", "Amount", "Removed"),
        reference_rows if reference_rows is not None else [("ab", "c", 10, "old"), ("a", "bc", 20, None), ("gone", "x", 7, None)], 4, 2)
    comparison_sheet = make_sheet("Current", "NewTable", ("Amount", "Code", "Part", "Extra"),
        comparison_rows if comparison_rows is not None else [(21, "bc", "a", None), (10, "c", "ab", "new"), (None, "y", "added", None)], 8, 5)

    def make_book(sheet, label):
        book = SimpleNamespace(Worksheets=_Collection(sheet), FullName="")
        book.Close = lambda **_: events.append(f"close {label}")

        def save(path, **options):
            assert options["FileFormat"] == 51
            book.FullName = path
            events.append(f"save {label}")

        book.SaveAs = save
        return book

    baseline, current = make_book(baseline_sheet, "baseline"), make_book(comparison_sheet, "comparison")

    def opened(_excel, source, *, read_only):
        if source == reference:
            assert read_only
            return baseline
        assert "close baseline" in events
        assert read_only if source == comparison else source.name.startswith(".ec-")
        return current

    monkeypatch.setattr(compare, "_excel_session", lambda: nullcontext(SimpleNamespace(Calculate=lambda: None)))
    monkeypatch.setattr(compare, "_open_workbook", opened)
    monkeypatch.setattr(compare, "_highlight", lambda sheet, cells: highlights.update({sheet.Name: cells}))
    return compare, reference, comparison, target, baseline, current, events, highlights


def test_key_comparison_matches_headers_and_reordered_composite_rows(tmp_path, monkeypatch):
    compare, reference, comparison, target, _baseline, _current, events, highlights = _key_books(tmp_path, monkeypatch)
    result = compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None,
        key_columns=("Part", "Code"), reference_table=("Baseline", "OldTable"), comparison_table=("Current", "NewTable"))
    assert result.changed_cells == 6
    assert (result.missing_rows, result.missing_columns, result.missing_sheets) == (1, ("Removed",), ())
    assert highlights == {"Current": [(8, 5), (9, 8), (10, 5), (10, 6), (10, 7), (10, 8)]}
    assert events == ["close baseline", "save comparison", "close comparison"]
    assert reference.read_bytes() == b"reference" and comparison.read_bytes() == b"comparison"
    assert target.read_bytes() == b"comparison"


@pytest.mark.parametrize("side", ["reference", "comparison"])
@pytest.mark.parametrize("duplicate", ["composite", "blank"])
def test_duplicate_keys_abort_before_highlight_and_publication(tmp_path, monkeypatch, side, duplicate):
    old = [(None, "", 1, None), ("", None, 2, None)] if duplicate == "blank" else [("a", "b", 1, None), ("a", "b", 2, None)]
    new = [(1, "", None, None), (2, None, "", None)] if duplicate == "blank" else [(1, "b", "a", None), (2, "b", "a", None)]
    compare, reference, comparison, target, _baseline, _current, events, highlights = _key_books(
        tmp_path, monkeypatch, reference_rows=old if side == "reference" else None, comparison_rows=new if side == "comparison" else None)
    with pytest.raises(WorkbookValidationError) as caught:
        compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None, key_columns=("Part", "Code"))
    message = str(caught.value)
    assert ("reference.xlsx" if side == "reference" else "comparison.xlsx") in message
    assert ("OldTable" if side == "reference" else "NewTable") in message
    assert ("4" in message and "5" in message) if side == "reference" else ("8" in message and "9" in message)
    assert not highlights and not target.exists() and not any(event.startswith("save") for event in events)
    assert set(tmp_path.iterdir()) == {reference, comparison}
    assert reference.read_bytes() == b"reference" and comparison.read_bytes() == b"comparison"


@pytest.mark.parametrize("side", ["reference", "comparison"])
@pytest.mark.parametrize("invalid", ["missing_table", "missing_key", "ambiguous_table"])
def test_key_mode_revalidates_selected_tables_and_headers(tmp_path, monkeypatch, side, invalid):
    compare, reference, comparison, target, baseline, current, _events, highlights = _key_books(tmp_path, monkeypatch)
    sheet = (baseline if side == "reference" else current).Worksheets.Item(1)
    selectors = {"reference_table": ("Baseline", "OldTable"), "comparison_table": ("Current", "NewTable")}
    if invalid == "missing_table":
        sheet.ListObjects.items.clear()
    elif invalid == "missing_key":
        for column in sheet.ListObjects.Item(1).ListColumns.items:
            if column.Name == "Code":
                column.Name = "Renamed"
    else:
        sheet.ListObjects.items.append(sheet.ListObjects.Item(1))
        selectors.pop(f"{side}_table")
    with pytest.raises(WorkbookValidationError):
        compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None,
            key_columns=("Part", "Code"), **selectors)
    assert not highlights and not target.exists()


def test_inspect_tables_reads_both_sources_sequentially_and_returns_headers(tmp_path, monkeypatch):
    compare, reference, comparison, _target, _baseline, _current, events, _highlights = _key_books(tmp_path, monkeypatch)
    tables = compare.CompareService().inspect_tables(reference, comparison)
    assert tables == (
        (compare.CompareTable("Baseline", "OldTable", ("Part", "Code", "Amount", "Removed")),),
        (compare.CompareTable("Current", "NewTable", ("Amount", "Code", "Part", "Extra")),),
    )
    assert events == ["close baseline", "close comparison"]


@pytest.mark.parametrize("side", ["reference", "comparison"])
def test_inspect_tables_rejects_a_source_without_a_table(tmp_path, monkeypatch, side):
    compare, reference, comparison, _target, baseline, current, _events, _highlights = _key_books(tmp_path, monkeypatch)
    (baseline if side == "reference" else current).Worksheets.Item(1).ListObjects.items.clear()
    with pytest.raises(WorkbookValidationError):
        compare.CompareService().inspect_tables(reference, comparison)


def test_key_values_distinguish_bool_text_errors_and_accept_numeric_equivalence(tmp_path, monkeypatch):
    old = [(True, "k", 1, None), (1, "k", 2, None), ("1", "k", 3, None), (-2146826281, "k", 4, None), (-2146826281.0, "k", 5, None)]
    new = [(5, "k", -2146826281.0, None), (4, "k", -2146826281, None), (3, "k", "1", None), (2, "k", 1.0, None), (1, "k", True, None)]
    compare, reference, comparison, target, _baseline, _current, _events, highlights = _key_books(tmp_path, monkeypatch, old, new)
    result = compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None, key_columns=("Part", "Code"))
    assert result.changed_cells == 0 and result.missing_rows == 0
    assert highlights == {"Current": []}


def test_table_selection_without_key_columns_is_rejected(tmp_path, monkeypatch):
    compare, reference, comparison, target = _setup(tmp_path, monkeypatch)
    with pytest.raises(WorkbookValidationError):
        compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None, reference_table=("Data", "Table1"))
    assert not target.exists()


def test_empty_tables_need_no_body_range_and_report_all_removed_keys(tmp_path, monkeypatch):
    compare, reference, comparison, target, _baseline, _current, _events, highlights = _key_books(tmp_path, monkeypatch, comparison_rows=[])
    result = compare.CompareService().execute(reference, comparison, target, progress=lambda *_: None, key_columns=("Part", "Code"))
    assert result.changed_cells == 0 and result.missing_rows == 3
    assert highlights == {"Current": []}
