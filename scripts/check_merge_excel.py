"""Opt-in native Excel check: python scripts/check_merge_excel.py."""

from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_splitter.excel_gateway import _excel_session, _open_workbook
from excel_splitter.file_signature import capture_signature
from excel_splitter.merge_service import MergeService


def main() -> None:
    with TemporaryDirectory(prefix="merge-excel-check-") as directory:
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
        before = tuple(capture_signature(source) for source in sources)
        service = MergeService()
        started = perf_counter()
        preview = service.preview(sources, root / "merged.xlsx")
        print(f"Preview: {perf_counter() - started:.3f}s for {len(sources)} files")
        assert preview.row_count == 5
        output = service.execute(preview, overwrite=False, progress=lambda *_: None)
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
        assert set(root.rglob("*.xlsx")) == {*sources, output}
    print("PASS: native Merge values, duplicates, formulas, formats, totals, filtered/hidden rows, styled blank expansion, outside content, and unchanged inputs")


if __name__ == "__main__":
    main()
