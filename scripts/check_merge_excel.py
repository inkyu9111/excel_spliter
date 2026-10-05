"""Opt-in native Excel check using the GUI's per-action worker lifetime."""

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from time import perf_counter
from contextlib import contextmanager

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_splitter.excel_gateway import _excel_session, _open_workbook
from excel_splitter.errors import SplitExecutionError, WorkbookValidationError
from excel_splitter.file_signature import capture_signature
from excel_splitter.merge_service import MergeService


def _gui_worker(action, *args, **kwargs):
    # The GUI starts a fresh thread for each preview/execution; do the same here.
    with ThreadPoolExecutor(max_workers=1) as worker:
        return worker.submit(action, *args, **kwargs).result()


def check_repeated_sessions() -> None:
    print("CHECK: retained COM objects across main/worker sessions and failures", flush=True)
    retained, failures = [], []

    def phase(count):
        for index in range(count):
            with _excel_session() as excel:
                book = excel.Workbooks.Add()
                sheet = book.Worksheets.Item(1)
                area = sheet.Range("A1:B2")
                area.Value2 = ((index, 2), (3, 4))
                assert area.Value2 == ((float(index), 2.0), (3.0, 4.0))
                retained.extend((excel, book, sheet, area))
                book.Close(SaveChanges=False)
            assert all(vars(proxy)["_oleobj_"] is None for proxy in retained)
        try:
            with _excel_session() as excel:
                book = excel.Workbooks.Add()
                area = book.Worksheets.Item(1).Range("A1")
                area.Value2 = "unsaved"
                retained.extend((excel, book, area))
                raise RuntimeError("deliberate lifecycle failure")
        except SplitExecutionError as exc:
            assert "deliberate lifecycle failure" in str(exc)
            failures.append(exc)  # Retain traceback locals beyond the apartment.
        assert all(vars(proxy)["_oleobj_"] is None for proxy in retained)

    phase(6)
    with ThreadPoolExecutor(max_workers=1) as worker:
        worker.submit(phase, 6).result()
    phase(2)


@contextmanager
def _new_table(excel, path, headers, rows):
    book = excel.Workbooks.Add()
    sheet = table = None
    try:
        book.Date1904 = False
        while book.Worksheets.Count > 1:
            book.Worksheets.Item(book.Worksheets.Count).Delete()
        sheet = book.Worksheets.Item(1)
        sheet.Name = "Data"
        width = len(headers)
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, width)).Value2 = (headers,)
        if rows:
            sheet.Range(sheet.Cells(2, 1), sheet.Cells(len(rows) + 1, width)).Value2 = rows
        table = sheet.ListObjects.Add(1, sheet.Range(sheet.Cells(1, 1),
            sheet.Cells(max(2, len(rows) + 1), width)), None, 1)
        table.Name = "DataTable"
        if not rows:
            table.DataBodyRange.Delete()
        yield book, sheet, table
        sheet.Calculate()
        book.SaveAs(str(path), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)
        book = sheet = table = None


def check_scalar_and_empty_tables(root: Path) -> None:
    print("CHECK: single-cell COM values, empty source order, all-empty Tables", flush=True)
    sources = tuple(root / name for name in (
        "scalar-a.xlsx", "scalar-b.xlsx", "empty-a.xlsx", "empty-b.xlsx",
        "empty-formula.xlsx", "precision-formula.xlsx",
    ))
    with _excel_session() as excel:
        for path, rows in zip(sources, ((("=literal",),), ((42,),), (), (), ((None,),), ((None,),))):
            with _new_table(excel, path, ("Value",), rows) as (book, sheet, table):
                if path == sources[0]:
                    table.DataBodyRange.NumberFormat = "@"
                    table.DataBodyRange.Value2 = "=literal"
                elif path == sources[4]:
                    table.DataBodyRange.Formula = '=IF(TRUE,"",1)'
                elif path == sources[5]:
                    table.DataBodyRange.Formula = "=1/3"
            book = sheet = table = None
    before = tuple(capture_signature(path) for path in sources)
    cases = (
        ((sources[0], sources[1]), (("=literal",), (42.0,))),
        ((sources[2], sources[0]), (("=literal",),)),
        ((sources[0], sources[2]), (("=literal",),)),
        ((sources[2], sources[3]), ()),
        ((sources[4], sources[5]), ((None,), (1 / 3,))),
    )
    service = MergeService()
    for index, (inputs, expected) in enumerate(cases):
        target = root / f"scalar-result-{index}.xlsx"
        preview = _gui_worker(service.preview, inputs, target)
        assert preview.row_count == len(expected)
        _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
        with _excel_session() as excel:
            book = _open_workbook(excel, target, read_only=True)
            try:
                sheet = book.Worksheets.Item(1)
                table = sheet.ListObjects.Item(1)
                assert table.ListRows.Count == len(expected)
                if expected:
                    actual = table.DataBodyRange.Value2
                    if len(expected) == 1:
                        actual = ((actual,),)
                    if inputs == sources[4:]:
                        # Excel may retain an empty string and round persisted
                        # numeric values to 15 significant digits.
                        assert actual[0] in ((None,), ("",)), actual
                        assert type(actual[1][0]) is float, actual
                        assert abs(actual[1][0] - 1 / 3) <= 1e-15, actual
                    else:
                        assert actual == expected, (actual, expected)
                    assert table.DataBodyRange.HasFormula is False
                else:
                    assert table.DataBodyRange is None
                assert book.LinkSources(1) is None
            finally:
                book.Close(SaveChanges=False)
                book = sheet = table = None
    assert tuple(capture_signature(path) for path in sources) == before


def check_types_and_failure_recovery(root: Path) -> None:
    print("CHECK: Excel errors, typed values, calculated columns, failed worker cleanup and retry", flush=True)
    sources = (root / "typed-a.xlsx", root / "typed-b.xlsx")
    headers = ("Name", "Boolean", "Number", "Text", "Error", "Literal", "Date", "Calculated")
    expected = (
        ("first", True, 1.25, "001", -2146826246, "=SUM(1,2)", 45000.0, 2.5),
        ("second", False, 0.0, "  text  ", -2146826281, "=SUM(1,2)", 45001.0, 0.0),
    )
    with _excel_session() as excel:
        for index, path in enumerate(sources):
            row = expected[index]
            with _new_table(excel, path, headers, (row,)) as (book, sheet, table):
                sheet.Range("D2").NumberFormat = "@"
                sheet.Range("D2").Value2 = row[3]
                sheet.Range("E2").Formula = "=NA()" if index == 0 else "=1/0"
                sheet.Range("F2").NumberFormat = "@"
                sheet.Range("F2").Value2 = row[5]
                sheet.Range("G2").NumberFormat = "yyyy-mm-dd"
                table.ListColumns.Item(8).DataBodyRange.Formula = "=[@Number]*2"
            book = sheet = table = None
    before = tuple(capture_signature(path) for path in sources)
    service = MergeService()
    target = root / "typed-result.xlsx"
    preview = _gui_worker(service.preview, sources, target)
    _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
    prior = capture_signature(target)
    overwrite = _gui_worker(service.preview, sources, target)

    def interrupt_after_first_source(completed, total, _label):
        if completed == 1 and total == len(sources):
            raise RuntimeError("synthetic-progress-failure")

    try:
        _gui_worker(service.execute, overwrite, overwrite=True, progress=interrupt_after_first_source)
    except SplitExecutionError as exc:
        assert "synthetic-progress-failure" in str(exc), str(exc)
    else:
        raise AssertionError("Interrupted merge reported success")
    assert capture_signature(target) == prior
    assert not tuple(root.glob(".em-*")), "Failed merge left temporary output"
    _gui_worker(service.execute, overwrite, overwrite=True, progress=lambda *_: None)
    with _excel_session() as excel:
        book = _open_workbook(excel, target, read_only=True)
        try:
            sheet = book.Worksheets.Item(1)
            table = sheet.ListObjects.Item(1)
            actual = table.DataBodyRange.Value2
            assert actual == expected, (actual, expected)
            assert all(type(row[1]) is bool for row in actual), actual
            assert all(type(row[4]) is int for row in actual), "Excel errors were converted to numbers"
            assert table.DataBodyRange.HasFormula is False
            assert sheet.Range("E2").Text == "#N/A" and sheet.Range("E3").Text == "#DIV/0!"
            assert sheet.Range("D2").NumberFormat == sheet.Range("F2").NumberFormat == "@"
            assert sheet.Range("G2:G3").NumberFormat == "yyyy-mm-dd"
        finally:
            book.Close(SaveChanges=False)
            book = sheet = table = None
    assert tuple(capture_signature(path) for path in sources) == before


def check_native_rejections(root: Path) -> None:
    print("CHECK: protected input, blocked Table expansion and corrupt workbook", flush=True)
    good = root / "valid-for-rejection.xlsx"
    protected = root / "protected.xlsx"
    blocked = root / "blocked-expansion.xlsx"
    corrupt = root / "corrupt.xlsx"
    with _excel_session() as excel:
        for path in (good, protected, blocked):
            with _new_table(excel, path, ("Value",), ((1,),)) as (book, sheet, table):
                if path == protected:
                    sheet.Protect(Password="synthetic-test")
                elif path == blocked:
                    sheet.Range("A3").Value2 = "must survive"
                    # Excel may auto-expand the Table when the adjacent cell
                    # is filled. Keep that value explicitly outside the Table.
                    table.Resize(sheet.Range("A1:A2"))
                    assert table.ListRows.Count == 1
                    assert sheet.Range("A3").Value2 == "must survive"
            book = sheet = table = None
        book = _open_workbook(excel, blocked, read_only=True)
        try:
            sheet = book.Worksheets.Item(1)
            table = sheet.ListObjects.Item(1)
            assert table.ListRows.Count == 1 and table.Range.Address == "$A$1:$A$2"
            assert sheet.Range("A3").Value2 == "must survive"
        finally:
            book.Close(SaveChanges=False)
            book = sheet = table = None
    corrupt.write_bytes(b"This is not an Excel workbook.")
    originals = {path: capture_signature(path) for path in (good, protected, blocked, corrupt)}
    service = MergeService()
    target = root / "rejected-output.xlsx"
    for inputs, message in (
        ((protected, good), "보호"),
        ((blocked, good), "확장 범위"),
        ((good, corrupt), None),
    ):
        try:
            _gui_worker(service.preview, inputs, target)
        except Exception as exc:
            if message:
                assert isinstance(exc, WorkbookValidationError), exc
                assert message in str(exc), str(exc)
            else:
                assert str(exc), "Corrupt input failed without an error message"
        else:
            raise AssertionError(f"Invalid input accepted: {inputs}")
        assert not target.exists()
        assert {path: capture_signature(path) for path in originals} == originals
        assert not tuple(root.glob(".em-*"))


def check_date_systems(root: Path) -> None:
    print("CHECK: 1904 date preservation and mixed-calendar rejection", flush=True)
    sources = tuple(root / f"calendar-{index}.xlsx" for index in range(3))
    with _excel_session() as excel:
        for index, path in enumerate(sources):
            with _new_table(excel, path, ("Date",), ((1000 + index,),)) as (book, sheet, table):
                book.Date1904 = index != 2
                sheet.Range("A2").NumberFormat = "yyyy-mm-dd"
                sheet.Columns.Item(1).ColumnWidth = 16
            book = sheet = table = None
    before = tuple(capture_signature(path) for path in sources)
    service = MergeService()
    target = root / "calendar-result.xlsx"
    preview = _gui_worker(service.preview, sources[:2], target)
    _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
    with _excel_session() as excel:
        book = _open_workbook(excel, target, read_only=True)
        try:
            sheet = book.Worksheets.Item(1)
            assert book.Date1904
            assert sheet.Range("A2:A3").Value2 == ((1000.0,), (1001.0,))
            assert sheet.Range("A2").Text == "1906-09-27"
            assert sheet.Range("A3").Text == "1906-09-28"
        finally:
            book.Close(SaveChanges=False)
            book = sheet = None
    prior = capture_signature(target)
    try:
        _gui_worker(service.preview, (sources[0], sources[2]), target)
    except WorkbookValidationError as exc:
        assert "날짜" in str(exc), str(exc)
    else:
        raise AssertionError("Different Excel date systems were accepted")
    assert capture_signature(target) == prior
    assert tuple(capture_signature(path) for path in sources) == before


def check_bulk_merge(root: Path) -> None:
    print("CHECK: 26,000-row Merge across bulk-read boundaries", flush=True)
    sources = (root / "bulk-a.xlsx", root / "bulk-b.xlsx")
    headers = ("ID", "Group", "Amount", "Zero", "Flag", "Text", "Blank", "Tail")
    batches = tuple(tuple(
        (index * 13000 + row, f"group-{row % 13}", row / 4, 0, row % 2 == 0, "001", None, row + 1)
        for row in range(13000)
    ) for index in range(2))
    with _excel_session() as excel:
        for path, rows in zip(sources, batches):
            with _new_table(excel, path, headers, rows) as (book, sheet, table):
                table.ListColumns.Item(3).DataBodyRange.NumberFormat = "0.000"
                table.ListColumns.Item(6).DataBodyRange.NumberFormat = "@"
                table.ListColumns.Item(6).DataBodyRange.Value2 = "001"
            book = sheet = table = None
    before = tuple(capture_signature(path) for path in sources)
    service = MergeService()
    target = root / "bulk-result.xlsx"
    started = perf_counter()
    preview = _gui_worker(service.preview, sources, target)
    preview_seconds = perf_counter() - started
    started = perf_counter()
    _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
    execution_seconds = perf_counter() - started
    with _excel_session() as excel:
        book = _open_workbook(excel, target, read_only=True)
        try:
            table = book.Worksheets.Item(1).ListObjects.Item(1)
            assert table.ListRows.Count == 26000
            assert table.DataBodyRange.Value2 == batches[0] + batches[1]
            assert table.DataBodyRange.HasFormula is False
            assert table.ListColumns.Item(3).DataBodyRange.NumberFormat == "0.000"
            assert table.ListColumns.Item(6).DataBodyRange.NumberFormat == "@"
        finally:
            book.Close(SaveChanges=False)
            book = table = None
    assert tuple(capture_signature(path) for path in sources) == before
    print(f"BENCHMARK: 2 x 13,000 rows x 8 columns; preview={preview_seconds:.3f}s; "
          f"execute with verification={execution_seconds:.3f}s; {26000 / execution_seconds:.0f} rows/s", flush=True)


def check_full_column_rule(root: Path) -> None:
    print("CHECK: one whole-column conditional format after 15-file Merge", flush=True)
    sources = tuple(root / f"cf-part-{index}.xlsx" for index in range(15))
    with _excel_session() as excel:
        for source in sources:
            book = excel.Workbooks.Add()
            try:
                while book.Worksheets.Count > 1:
                    book.Worksheets.Item(book.Worksheets.Count).Delete()
                sheet = book.Worksheets.Item(1)
                sheet.Name = "Data"
                sheet.Range("A1:H1").Value2 = (tuple(f"Column{i}" for i in range(1, 9)),)
                sheet.Range("A2:H3").Value2 = (tuple(range(1, 9)), tuple(range(11, 19)))
                sheet.ListObjects.Add(1, sheet.Range("A1:H3"), None, 1).Name = "DataTable"
                rule = sheet.Range("H:H").FormatConditions.Add(2, 3, "=$H1>0", "")
                rule.StopIfTrue = True
                rule.Font.Bold = True
                rule.Font.Color = 255
                rule.Interior.Color = 65535
                rule.NumberFormat = "0.0000"
                book.SaveAs(str(source), FileFormat=51)
            finally:
                book.Close(SaveChanges=False)
                # Release child COM proxies before their Excel session ends.
                book = sheet = rule = None
    before = tuple(capture_signature(source) for source in sources)
    service = MergeService()
    preview = _gui_worker(service.preview, sources, root / "cf-merged.xlsx")
    output = _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
    assert tuple(capture_signature(source) for source in sources) == before
    with _excel_session() as excel:
        book = _open_workbook(excel, output, read_only=True)
        try:
            sheet = book.Worksheets.Item(1)
            assert sheet.ListObjects.Item(1).ListRows.Count == 30
            assert sheet.Cells.FormatConditions.Count == 1
            rule = sheet.Cells.FormatConditions.Item(1)
            assert rule.AppliesTo.Address == "$H:$H"
            assert rule.Formula1 == "=$H1>0"
            assert rule.Priority == 1 and rule.StopIfTrue
            assert rule.Font.Bold and rule.Font.Color == 255
            assert rule.Interior.Color == 65535 and rule.NumberFormat == "0.0000"
            assert sheet.Range("H31").DisplayFormat.Interior.Color == 65535
            assert book.LinkSources(1) is None
        finally:
            book.Close(SaveChanges=False)
            book = sheet = rule = None


def check_filtered_template(root: Path, sources: tuple[Path, ...]) -> Path:
    print("CHECK: actively filtered first template includes every source row", flush=True)
    before = tuple(capture_signature(path) for path in sources)
    service = MergeService()
    target = root / "filtered-template-result.xlsx"
    preview = _gui_worker(service.preview, tuple(reversed(sources)), target)
    assert preview.row_count == 5
    _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
    with _excel_session() as excel:
        book = _open_workbook(excel, target, read_only=True)
        try:
            sheet = book.Worksheets.Item(1)
            table = sheet.ListObjects.Item(1)
            assert table.ListRows.Count == 5
            assert table.DataBodyRange.Value2 == (
                ("repeat", 4.0, "=literal"), ("filtered", 5.0, "=literal"),
                ("hidden", 6.0, "=literal"), ("repeat", 4.0, "=literal"),
                ("repeat", 4.0, "=literal"),
            )
            assert table.AutoFilter.FilterMode is False
            assert table.DataBodyRange.HasFormula is False
            assert table.ShowTotals is False
            assert table.DataBodyRange.Cells(1, 2).NumberFormat == "0.000"
            assert table.DataBodyRange.Cells(4, 2).NumberFormat == "0.00"
            assert sheet.Range("A1").Value2 == "Ignored layout"
            assert abs(sheet.Columns.Item(2).ColumnWidth - 12) < 0.1
            assert book.LinkSources(1) is None
        finally:
            book.Close(SaveChanges=False)
            book = sheet = table = None
    assert tuple(capture_signature(path) for path in sources) == before
    return target


def main() -> None:
    check_repeated_sessions()
    scratch = Path(__file__).resolve().parents[1] / "build"
    scratch.mkdir(exist_ok=True)
    with TemporaryDirectory(prefix="merge-excel-check-", dir=scratch) as directory:
        root = Path(directory)
        sources = (root / "first" / "part.xlsx", root / "second" / "part.xlsx")
        with _excel_session() as excel:
            for index, source in enumerate(sources):
                source.parent.mkdir()
                book = excel.Workbooks.Add()
                try:
                    while book.Worksheets.Count > 1:
                        book.Worksheets.Item(book.Worksheets.Count).Delete()
                    sheet = book.Worksheets.Item(1)
                    sheet.Name = "Data"
                    sheet.Range("A1").Value2 = "First layout" if index == 0 else "Ignored layout"
                    sheet.Columns.Item(2).ColumnWidth = 23 if index == 0 else 12
                    rows = [("repeat", 4, "literal"), ("repeat", 4, "literal")] if index == 0 else [
                        ("repeat", 4, "literal"), ("filtered", 5, "literal"), ("hidden", 6, "literal")
                    ]
                    sheet.Range("B5:D5").Value2 = (("Name", "Amount", "Literal"),)
                    sheet.Range(f"B6:D{5 + len(rows)}").Value2 = tuple(rows)
                    sheet.Range("C6").Formula = "=2+2"
                    table = sheet.ListObjects.Add(1, sheet.Range(f"B5:D{5 + len(rows)}"), None, 1)
                    table.Name = "DataTable"
                    table.TableStyle = "TableStyleMedium2"
                    table.ListColumns.Item(2).DataBodyRange.NumberFormat = "0.00" if index == 0 else "0.000"
                    table.ListColumns.Item(3).DataBodyRange.NumberFormat = "@"
                    table.ListColumns.Item(3).DataBodyRange.Value2 = "=literal"
                    if index == 0:
                        table.ShowTotals = True
                        table.ListColumns.Item(2).TotalsCalculation = 1
                        # Split outputs can retain styled, empty cells below the Table.
                        residual_style = book.Styles.Add("MergeResidualStyle")
                        residual_style.NumberFormat = "0.00"
                        sheet.Range("B9:D52").Style = residual_style.Name
                        sheet.Range("C10").Style = book.Styles.Item(1).Name
                        sheet.Range("B14").Value2 = "Keep below merged Table"
                    else:
                        table.Range.AutoFilter(Field=1, Criteria1="repeat")
                        sheet.Rows.Item(6).Hidden = True
                    sheet.Calculate()
                    book.SaveAs(str(source), FileFormat=51)
                finally:
                    book.Close(SaveChanges=False)
                    book = sheet = table = residual_style = None
        before = tuple(capture_signature(source) for source in sources)
        service = MergeService()
        started = perf_counter()
        preview = _gui_worker(service.preview, sources, root / "merged.xlsx")
        print(f"Preview: {perf_counter() - started:.3f}s for {len(sources)} files")
        assert preview.row_count == 5
        output = _gui_worker(service.execute, preview, overwrite=False, progress=lambda *_: None)
        assert tuple(capture_signature(source) for source in sources) == before
        with _excel_session() as excel:
            book = _open_workbook(excel, output, read_only=True)
            try:
                sheet = book.Worksheets.Item(1)
                table = sheet.ListObjects.Item(1)
                assert table.ListRows.Count == 5
                assert (table.Range.Row, table.Range.Column, table.Range.Rows.Count, table.Range.Columns.Count) == (5, 2, 7, 3)
                assert (table.DataBodyRange.Row, table.DataBodyRange.Column, table.DataBodyRange.Rows.Count) == (6, 2, 5)
                assert table.DataBodyRange.Value2 == (
                    ("repeat", 4.0, "=literal"), ("repeat", 4.0, "=literal"),
                    ("repeat", 4.0, "=literal"), ("filtered", 5.0, "=literal"),
                    ("hidden", 6.0, "=literal"),
                )
                assert table.DataBodyRange.HasFormula is False
                assert table.ShowTotals and table.TotalsRowRange.Cells(1, 2).Value2 == 23.0
                assert table.DataBodyRange.Cells(1, 2).NumberFormat == "0.00"
                assert table.DataBodyRange.Cells(3, 2).NumberFormat == "0.000"
                assert sheet.Range("A1").Value2 == "First layout"
                assert sheet.Range("B14").Value2 == "Keep below merged Table"
                assert abs(sheet.Columns.Item(2).ColumnWidth - 23) < 0.1
                assert book.LinkSources(1) is None
            finally:
                book.Close(SaveChanges=False)
                book = sheet = table = None
        filtered_output = check_filtered_template(root, sources)
        assert set(root.rglob("*.xlsx")) == {*sources, output, filtered_output}
        check_full_column_rule(root)
        check_scalar_and_empty_tables(root)
        check_types_and_failure_recovery(root)
        check_native_rejections(root)
        check_date_systems(root)
        check_bulk_merge(root)
    print("PASS: native Merge data/types/errors, scalar/empty Tables, duplicates, calculated columns, formats/totals, filtered/hidden rows, expansion, conditional formats, failed worker recovery, validation and unchanged inputs", flush=True)


if __name__ == "__main__":
    main()
