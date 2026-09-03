"""Opt-in synthetic native Excel check: python scripts/check_compare_excel.py."""

from pathlib import Path
import sys
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.excel_gateway import _excel_session, _open_workbook
from excel_splitter.file_signature import capture_signature


def _check_key_tables(root: Path) -> None:
    from excel_splitter.compare_service import CompareService

    sources = (root / "key-baseline.xlsx", root / "key-comparison.xlsx")
    with _excel_session() as excel:
        for index, source in enumerate(sources):
            book = excel.Workbooks.Add()
            try:
                while book.Worksheets.Count > 1:
                    book.Worksheets.Item(book.Worksheets.Count).Delete()
                sheet = book.Worksheets.Item(1)
                sheet.Name = "Baseline" if index == 0 else "Current"
                address = "B3:E6" if index == 0 else "E7:H10"
                sheet.Range(address).Value2 = (
                    (("Part", "Code", "Amount", "Removed"), ("ab", "c", 10, "old"), ("a", "bc", 20, None), ("gone", "x", 7, None))
                    if index == 0 else
                    (("Amount", "Code", "Part", "Extra"), (21, "bc", "a", None), (10, "c", "ab", "new"), (None, "y", "added", None))
                )
                if index:
                    sheet.Range("E8").Formula = "=20+1"
                    sheet.Range("E9").Formula = "=5+5"
                    sheet.Range("E8:E10").NumberFormat = "0.00"
                table = sheet.ListObjects.Add(1, sheet.Range(address), None, 1)
                table.Name = "OldTable" if index == 0 else "NewTable"
                if index:
                    table.ShowTotals = True
                    sheet.Range("K2:L3").Value2 = (("Other", "Value"), ("keep", 99))
                    sheet.ListObjects.Add(1, sheet.Range("K2:L3"), None, 1).Name = "SideTable"
                    sheet.Range("A1").Value2 = "outside content"
                book.SaveAs(str(source), FileFormat=51)
            finally:
                book.Close(SaveChanges=False)
    before = tuple(capture_signature(source) for source in sources)
    service = CompareService()
    tables = service.inspect_tables(*sources)
    assert len(tables[0]) == 1 and len(tables[1]) == 2
    target = root / "key-result.xlsx"
    result = service.execute(*sources, target, progress=lambda *_: None, key_columns=("Part", "Code"),
        reference_table=("Baseline", "OldTable"), comparison_table=("Current", "NewTable"))
    assert (result.changed_cells, result.missing_rows, result.missing_columns, result.missing_sheets) == (6, 1, ("Removed",), ())
    assert tuple(capture_signature(source) for source in sources) == before
    with _excel_session() as excel:
        book = _open_workbook(excel, target, read_only=True)
        try:
            sheet = book.Worksheets.Item("Current")
            for address in ("E8", "H9", "E10", "F10", "G10", "H10"):
                assert sheet.Range(address).Interior.Color == 65535, address
            for address in ("E7", "E9", "F8", "G8", "E11", "K3"):
                assert sheet.Range(address).Interior.Color != 65535, address
            assert sheet.Range("E8").Formula == "=20+1" and sheet.Range("E9").Formula == "=5+5"
            assert sheet.Range("E10").Value2 is None
            assert sheet.Range("E8").NumberFormat == "0.00"
            assert sheet.Range("F8:G10").Value2 == (("bc", "a"), ("c", "ab"), ("y", "added"))
            assert sheet.ListObjects.Item("NewTable").ShowTotals
            assert sheet.ListObjects.Item("SideTable").DataBodyRange.Value2 == (("keep", 99.0),)
            assert sheet.Range("A1").Value2 == "outside content"
        finally:
            book.Close(SaveChanges=False)
    assert tuple(capture_signature(source) for source in sources) == before


def main() -> None:
    from excel_splitter.compare_service import CompareService

    scratch = Path(__file__).resolve().parents[1] / "build"
    scratch.mkdir(exist_ok=True)
    with TemporaryDirectory(prefix="compare-excel-check-", dir=scratch) as directory:
        root = Path(directory)
        sources = (root / "baseline" / "book.xlsx", root / "comparison" / "book.xlsx")
        with _excel_session() as excel:
            for index, source in enumerate(sources):
                source.parent.mkdir()
                book = excel.Workbooks.Add()
                try:
                    while book.Worksheets.Count > 1:
                        book.Worksheets.Item(book.Worksheets.Count).Delete()
                    sheet = book.Worksheets.Item(1)
                    sheet.Name = "Data"
                    sheet.Range("A1:D3").Value2 = (
                        ("same", 1 if index == 0 else 2, True if index == 0 else 1, "x" if index == 0 else " x"),
                        ("removed" if index == 0 else None, None if index == 0 else "added", "text", 4),
                        (None, None, None, None),
                    )
                    sheet.Range("C2").NumberFormat = "@"
                    sheet.Range("C2").Value2 = "1" if index == 0 else "=literal"
                    sheet.Range("D2").Formula = "=2+2" if index == 0 else "=1+3"
                    sheet.Range("A3").Formula = "=1" if index == 0 else "=2"
                    sheet.Range("B3").Formula = '=""'
                    sheet.Range("C3").NumberFormat = "@" if index == 0 else "General"
                    sheet.Range("C3").Value2 = "7" if index == 0 else 7
                    if index == 0:
                        sheet.Range("D3").Formula = "=1/0"
                    else:
                        sheet.Range("D3").Value2 = -2146826281.0
                    if index == 0:
                        sheet.Range("F8").Value2 = "removed outside comparison range"
                        sheet.Range("A33000").Value2 = "removed after chunk boundary"
                    sheet.Range("A1").Interior.Color = 255
                    sheet.Range("B1").Interior.Color = 65280
                    rule = sheet.Range("B1").FormatConditions.Add(2, 3, "=TRUE", "")
                    rule.Interior.Color = 255
                    rule.Font.Bold = True
                    sheet.Range("A1").Font.Bold = True
                    sheet.Columns.Item(1).ColumnWidth = 24
                    sheet.Rows.Item(2).Hidden = True
                    extra = book.Worksheets.Add(After=sheet)
                    extra.Name = "Missing" if index == 0 else "Added"
                    extra.Range("B2").Value2 = "extra sheet"
                    extra.Range("D4").Formula = '=""'
                    untouched = book.Worksheets.Add(After=extra)
                    untouched.Name = "Untouched"
                    untouched.Visible = 0
                    untouched.Range("A1").Value2 = "keep"
                    sheet.Calculate()
                    book.SaveAs(str(source), FileFormat=51)
                finally:
                    book.Close(SaveChanges=False)
        before = tuple(capture_signature(source) for source in sources)
        output = root / "result.xlsx"
        result = CompareService().execute(*sources, output, progress=lambda *_: None)
        assert result.changed_cells == 12, result
        assert result.missing_sheets == ("Missing",), result
        assert tuple(capture_signature(source) for source in sources) == before
        with _excel_session() as excel:
            book = _open_workbook(excel, output, read_only=True)
            try:
                assert book.Worksheets.Count == 3
                sheet = book.Worksheets.Item("Data")
                for address in ("B1", "C1", "D1", "A2", "B2", "C2", "A3", "C3", "D3", "F8", "A33000"):
                    assert sheet.Range(address).Interior.Color == 65535, address
                assert sheet.Range("F8").Value2 is None
                assert sheet.Range("A33000").Value2 is None
                assert sheet.Range("A1").Interior.Color == 255
                assert sheet.Range("B1").DisplayFormat.Interior.Color == 65535
                assert sheet.Range("B1").DisplayFormat.Font.Bold
                assert sheet.Range("B1").FormatConditions.Count == 2
                assert sheet.Range("D2").Interior.Color != 65535
                assert sheet.Range("B3").Interior.Color != 65535
                assert sheet.Range("D2").Formula == "=1+3"
                assert sheet.Range("A3").Formula == "=2"
                assert sheet.Range("C2").Value2 == "=literal"
                assert sheet.Range("C2").NumberFormat == "@"
                assert sheet.Range("A1").Font.Bold
                assert sheet.Rows.Item(2).Hidden
                assert abs(sheet.Columns.Item(1).ColumnWidth - 24) < 0.1
                assert book.Worksheets.Item("Added").Range("B2").Interior.Color == 65535
                assert book.Worksheets.Item("Added").Range("D4").Interior.Color != 65535
                assert book.Worksheets.Item("Untouched").Visible == 0
            finally:
                book.Close(SaveChanges=False)
        existing_signature = capture_signature(output)
        try:
            CompareService().execute(*sources, output, progress=lambda *_: None)
        except WorkbookValidationError:
            pass
        else:
            raise AssertionError("Existing output was accepted")
        assert capture_signature(output) == existing_signature
        same = CompareService().execute(sources[0], sources[0], root / "same.xlsx", progress=lambda *_: None)
        assert same.changed_cells == 0 and same.missing_sheets == ()
        assert tuple(capture_signature(source) for source in sources) == before
        assert set(root.rglob("*.xlsx")) == {*sources, output, same.target}
        _check_key_tables(root)
    print("PASS: native Compare positions and composite keys, reordered rows/columns, yellow fills, blank/extra cells, sheet differences, formulas/styles, types, chunk boundary, same filenames, same source, unchanged inputs, and existing-output rejection")


if __name__ == "__main__":
    main()
