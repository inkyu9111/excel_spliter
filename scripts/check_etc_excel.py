"""Opt-in synthetic native Excel check: python scripts/check_etc_excel.py."""

from pathlib import Path
import sys
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.etc_service import EtcService
from excel_splitter.excel_artifacts import threaded_comment_count, unsupported_threaded_comments_error
from excel_splitter.excel_gateway import _excel_session, _open_workbook
from excel_splitter.file_signature import capture_signature


def _counts(sheet):
    return int(sheet.Shapes.Count), int(sheet.Comments.Count), threaded_comment_count(sheet)


def main() -> None:
    scratch = Path(__file__).resolve().parents[1] / "build"
    scratch.mkdir(exist_ok=True)
    with TemporaryDirectory(prefix="etc-excel-check-", dir=scratch) as directory:
        root = Path(directory)
        source = root / "source.xlsx"
        styled_cells = ("B3", "B4", "B5", "B8")  # header, both bands, totals
        with _excel_session() as excel:
            book = excel.Workbooks.Add()
            try:
                while book.Worksheets.Count > 1:
                    book.Worksheets.Item(book.Worksheets.Count).Delete()
                selected = book.Worksheets.Item(1)
                selected.Name = "Selected"
                other = book.Worksheets.Add(After=selected)
                other.Name = "Other"
                for sheet in (selected, other):
                    sheet.Range("B3:D7").Value2 = (
                        ("Item", "Amount", "Note"), ("a", 20, "keep"),
                        ("b", 30, "keep"), ("c", 40, "keep"), ("d", 50, "keep"),
                    )
                    sheet.Range("C4").Formula = "=10*2"
                    table = sheet.ListObjects.Add(1, sheet.Range("B3:D7"), None, 1)
                    table.Name = f"{sheet.Name}Table"
                    table.TableStyle = "TableStyleMedium2"
                    table.ShowTableStyleRowStripes = True
                    table.ShowTotals = True
                    table.ListColumns.Item(2).TotalsCalculation = 1
                    sheet.Range("C4:C7").NumberFormat = "0.00"
                    sheet.Range("D4").Font.Bold = True
                    sheet.Range("D4").Borders.Item(9).LineStyle = 1  # bottom edge
                    sheet.Range("G2").Value2 = "conditional fill"
                    condition = sheet.Range("G2").FormatConditions.Add(Type=2, Formula1="=TRUE")
                    condition.Interior.Color = 255
                    sheet.Range("D5").AddComment("legacy note")
                    try:
                        sheet.Range("D6").AddCommentThreaded("threaded comment")
                    except Exception as exc:
                        if not unsupported_threaded_comments_error(exc):
                            raise
                    sheet.Shapes.AddShape(1, 300, 100, 80, 40)
                    sheet.Calculate()
                table_colors = {address: selected.Range(address).DisplayFormat.Interior.Color for address in styled_cells}
                for sheet in (selected, other):
                    for address in (*styled_cells, "G2", "Z100000"):
                        sheet.Range(address).Interior.Color = 65535
                before_counts = _counts(selected)
                other_counts = _counts(other)
                book.SaveAs(str(source), FileFormat=51)
            finally:
                book.Close(SaveChanges=False)
        signature = capture_signature(source)
        service = EtcService()
        assert service.inspect_source(source) == ("Selected", "Other")
        outputs = []
        for remove_artifacts, reset_fill in ((True, False), (False, True), (True, True)):
            output = root / f"result-{int(remove_artifacts)}-{int(reset_fill)}.xlsx"
            service.execute(source, "Selected", output, remove_artifacts=remove_artifacts,
                            reset_fill=reset_fill, progress=lambda *_: None)
            outputs.append(output)
            assert capture_signature(source) == signature
            with _excel_session() as excel:
                book = _open_workbook(excel, output, read_only=True)
                try:
                    sheet = book.Worksheets.Item("Selected")
                    table = sheet.ListObjects.Item("SelectedTable")
                    assert _counts(sheet) == ((0, 0, 0) if remove_artifacts else before_counts)
                    style = table.TableStyle
                    assert str(getattr(style, "Name", style)) == "TableStyleMedium2"
                    assert table.ShowTableStyleRowStripes and table.ShowTotals
                    for address in styled_cells:
                        if reset_fill:
                            assert sheet.Range(address).DisplayFormat.Interior.Color == table_colors[address], address
                        else:
                            assert sheet.Range(address).Interior.Color == 65535, address
                    assert sheet.Range("Z100000").Interior.Pattern == (-4142 if reset_fill else 1)
                    assert sheet.Range("Z100000").Value2 is None
                    assert sheet.Range("G2").FormatConditions.Count == 1
                    assert sheet.Range("G2").DisplayFormat.Interior.Color == 255
                    assert sheet.Range("C4").Formula == "=10*2" and sheet.Range("C4").Value2 == 20
                    assert sheet.Range("C4").NumberFormat == "0.00"
                    assert sheet.Range("D4").Font.Bold and sheet.Range("D4").Borders.Item(9).LineStyle == 1
                    assert sheet.Range("B4:B7").Value2 == (("a",), ("b",), ("c",), ("d",))
                    assert table.TotalsRowRange.Cells(1, 2).Value2 == 140
                    other = book.Worksheets.Item("Other")
                    assert _counts(other) == other_counts
                    assert other.Range("B3").Interior.Color == 65535
                    assert other.Range("Z100000").Interior.Color == 65535
                    assert other.Range("G2").FormatConditions.Count == 1
                    assert other.Range("C4").Formula == "=10*2"
                finally:
                    book.Close(SaveChanges=False)
        prior = capture_signature(outputs[-1])
        try:
            service.execute(source, "Selected", outputs[-1], remove_artifacts=True, reset_fill=True, progress=lambda *_: None)
        except WorkbookValidationError:
            pass
        else:
            raise AssertionError("Existing output was accepted")
        assert capture_signature(outputs[-1]) == prior and capture_signature(source) == signature
        assert set(root.iterdir()) == {source, *outputs}
    print("PASS: native Etc options, shapes/notes/supported threaded comments, direct fill reset, Table headers/bands/totals, conditional fill, distant blank cells, formulas/styles, other sheet, original hashes, and output rejection")


if __name__ == "__main__":
    main()
