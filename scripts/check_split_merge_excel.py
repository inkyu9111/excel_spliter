"""Opt-in native Excel regression: python scripts/check_split_merge_excel.py.

Uses disposable synthetic workbooks only. Run sequentially with other native
checks because Excel's clipboard is shared. Requires desktop Excel and pywin32.
"""

from collections import Counter
from contextlib import contextmanager
from pathlib import Path
import shutil
import sys
from tempfile import TemporaryDirectory
from xml.etree import ElementTree
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.excel_gateway import ExcelComGateway, _excel_session, _open_workbook
from excel_splitter.file_signature import capture_signature
from excel_splitter.merge_service import MergeService
from excel_splitter.parallel_writer import ParallelWriteAborted
from excel_splitter.split_service import SplitService


HEADERS = ("지역", "순번", "금액", "계산", "문자")
ROWS = (
    ("서울", 1, 10, 20, "=literal"),
    ("부산", 2, 7, 14, "=literal"),
    ("서울", 1, 10, 20, "=literal"),  # An exact duplicate must survive.
    (None, 3, 5, 10, "=literal"),
    ("부산", 4, 9, 18, "=literal"),
    ("서울", 5, 12, 24, "=literal"),
)
# Fixed expectations, independent of the production classifier and row indexes.
GROUPS = {
    "서울": (ROWS[0], ROWS[2], ROWS[5]),
    "부산": (ROWS[1], ROWS[4]),
    "": (ROWS[3],),
}
MERGED = GROUPS["부산"] + GROUPS[""] + GROUPS["서울"]


def checkpoint(name):
    print(f"CHECK: {name}", flush=True)


@contextmanager
def rejected(error_type, message):
    try:
        yield
    except error_type as exc:
        assert message in str(exc), (message, str(exc))
    else:
        raise AssertionError(f"Expected {error_type.__name__}: {message}")


def make_source(excel, path, *, empty=False, no_table=False, mismatch=False, multisheet=False, filtered=False):
    book = excel.Workbooks.Add()
    try:
        while book.Worksheets.Count > 1:
            book.Worksheets.Item(book.Worksheets.Count).Delete()
        sheet = book.Worksheets.Item(1)
        sheet.Name = "분류 표"
        if not no_table:
            sheet.Range("A1").Value2 = "합성 회귀 자료"
            sheet.Columns.Item(2).ColumnWidth = 25
            headers = ("다른 지역", *HEADERS[1:]) if mismatch else HEADERS
            sheet.Range("B5:F5").Value2 = (headers,)
            sheet.Range("B6:F11").Value2 = ROWS
            table = sheet.ListObjects.Add(1, sheet.Range("B5:F11"), None, 1)
            table.Name = "SyntheticData"
            table.TableStyle = "TableStyleMedium2"
            table.ListColumns.Item(3).DataBodyRange.NumberFormat = "0.00"
            table.ListColumns.Item(3).DataBodyRange.Font.Bold = True
            table.ListColumns.Item(3).DataBodyRange.Font.Color = 255
            table.ListColumns.Item(4).DataBodyRange.Formula = "=[@금액]*2"
            table.ListColumns.Item(5).DataBodyRange.NumberFormat = "@"
            table.ListColumns.Item(5).DataBodyRange.Value2 = "=literal"
            if empty:
                table.DataBodyRange.Delete()
                assert table.ListRows.Count == 0
            elif multisheet:
                # Manually hidden input rows still belong to the positive case.
                sheet.Rows.Item(7).Hidden = True
            if filtered:
                table.Range.AutoFilter(Field=1, Criteria1="서울")
        if multisheet:
            other = book.Worksheets.Add(After=sheet)
            other.Name = "참조 시트"
            other.Range("A1").Value2 = 99
            sheet.Range("H1").Formula = "='참조 시트'!A1"
        sheet.Calculate()
        book.SaveAs(str(path), FileFormat=51)
    finally:
        book.Close(SaveChanges=False)


def check_output(path, expected, *, formulas, broken_reference=False):
    checkpoint(f"verify output {path.name}")
    with _excel_session() as excel:
        book = _open_workbook(excel, path, read_only=True)
        try:
            assert book.Sheets.Count == book.Worksheets.Count == 1
            sheet = book.Worksheets.Item(1)
            assert sheet.Name == "분류 표"
            assert sheet.ListObjects.Count == 1
            table = sheet.ListObjects.Item(1)
            assert table.HeaderRowRange.Value2 == (HEADERS,)
            assert table.ListRows.Count == len(expected)
            actual = table.DataBodyRange.Value2 if expected else ()
            assert actual == expected, (path.name, actual, expected)
            assert str(table.TableStyle) == "TableStyleMedium2"
            assert sheet.Range("A1").Value2 == "합성 회귀 자료"
            assert abs(sheet.Columns.Item(2).ColumnWidth - 25) < 0.1
            if expected:
                amount = table.ListColumns.Item(3).DataBodyRange
                assert amount.NumberFormat == "0.00"
                assert amount.Font.Bold and amount.Font.Color == 255
                assert bool(table.ListColumns.Item(4).DataBodyRange.HasFormula) == formulas
                assert table.ListColumns.Item(5).DataBodyRange.HasFormula is False
                if not formulas:
                    assert table.DataBodyRange.HasFormula is False
            if broken_reference:
                # Documented limitation: Split removes other sheets, not references.
                assert sheet.Range("H1").HasFormula
                assert "#REF!" in sheet.Range("H1").Formula
            assert book.LinkSources(1) is None
        finally:
            book.Close(SaveChanges=False)


def main():
    import win32con
    import win32file

    with TemporaryDirectory(prefix="split-merge-check-") as directory:
        root = Path(directory) / "한글 경로 공백"
        root.mkdir()
        source = root / "분할 원본.xlsx"
        blank = root / "빈 시트.xlsx"
        empty = root / "빈 Table.xlsx"
        mismatch = root / "다른 머리글.xlsx"
        filtered = root / "활성 필터 원본.xlsx"
        checkpoint("create synthetic input workbooks")
        with _excel_session() as excel:
            make_source(excel, source, multisheet=True)
            make_source(excel, blank, no_table=True)
            make_source(excel, empty, empty=True)
            make_source(excel, mismatch, mismatch=True)
            make_source(excel, filtered, multisheet=True, filtered=True)
        originals = {path: capture_signature(path) for path in (source, blank, empty, mismatch, filtered)}
        split = SplitService(ExcelComGateway())
        merge = MergeService()
        try:
            checkpoint("verify persisted fixture sheets and Split sheet discovery")
            with ZipFile(source) as package:
                workbook_xml = ElementTree.fromstring(package.read("xl/workbook.xml"))
            stored_names = tuple(item.attrib["name"] for item in workbook_xml.findall(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/"
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
            ))
            assert Counter(stored_names) == Counter(("분류 표", "참조 시트")), (
                "Synthetic fixture has unexpected persisted sheets", stored_names,
            )
            actual_names = split.list_sheets(source)
            print(f"SHEETS: persisted={stored_names!r}; Excel={actual_names!r}", flush=True)
            assert Counter(actual_names) == Counter(stored_names), (
                "Split sheet discovery differs from persisted fixture", actual_names, stored_names,
            )
            checkpoint("reject blank worksheet and zero-row Split Table")
            for path, message in ((blank, "Table이 없습니다"), (empty, "데이터 행이 없습니다")):
                with rejected(WorkbookValidationError, message):
                    split.preview(path, "분류 표", "지역", "% 분할", root)
            checkpoint("Split preview and native execution")
            preview = split.preview(source, "분류 표", "지역", "% 분할", root)
            assert preview.snapshot.row_count == 6
            assert {group.label: group.count for group in preview.snapshot.groups} == {
                "서울": 3, "부산": 2, "": 1,
            }
            result = split.execute(preview, overwrite=False, progress=lambda *_: None)
            assert not result.failed, result.failed
            outputs = {target.label: target.path for target in preview.targets}
            assert set(result.succeeded) == set(outputs.values())
            for label, path in outputs.items():
                check_output(path, GROUPS[label], formulas=True, broken_reference=True)
            print("PASS: Split rows, duplicates, blank group, formulas, styles, selected sheet and known #REF limitation", flush=True)

            # Explicitly choose a different group order: Merge is concatenation,
            # not reconstruction of the interleaved original row order.
            inputs = (outputs["부산"], outputs[""], outputs["서울"])
            target = root / "병합 결과.xlsx"
            signatures = {path: capture_signature(path) for path in inputs}
            checkpoint("reject duplicate inputs, output as input, mismatched headers and unsupported sheets")
            for sources, destination, message in (
                ((inputs[0], inputs[0]), target, "중복 선택"),
                (inputs, inputs[0], "원본 파일을 덮어쓸 수 없습니다"),
                ((inputs[0], mismatch), target, "열 이름과 순서"),
                ((inputs[0], source), target, "워크시트가 정확히 하나"),
                ((inputs[0], blank), target, "Table이 없습니다"),
            ):
                with rejected(WorkbookValidationError, message):
                    merge.preview(sources, destination)
            checkpoint("merge Split outputs in explicit group order")
            merged_preview = merge.preview(inputs, target)
            assert merged_preview.row_count == 6
            merge.execute(merged_preview, overwrite=False, progress=lambda *_: None)
            check_output(target, MERGED, formulas=False, broken_reference=True)
            assert Counter(MERGED) == Counter(ROWS)
            assert all(capture_signature(path) == signature for path, signature in signatures.items())

            # An empty Table is a valid zero-row Merge input, including as template.
            checkpoint("merge with a zero-row Table as template")
            empty_target = root / "빈 표 포함 병합.xlsx"
            empty_preview = merge.preview((empty, inputs[0]), empty_target)
            assert empty_preview.row_count == 2
            merge.execute(empty_preview, overwrite=False, progress=lambda *_: None)
            check_output(empty_target, GROUPS["부산"], formulas=False)

            checkpoint("Merge overwrite approval and locked output preservation")
            overwrite_preview = merge.preview(inputs, target)
            before_target = capture_signature(target)
            with rejected(WorkbookValidationError, "덮어쓰기 승인"):
                merge.execute(overwrite_preview, overwrite=False, progress=lambda *_: None)
            assert capture_signature(target) == before_target
            # Native deny-delete handle models an open/locked output while still
            # allowing the application's normal read-only signature verification.
            handle = win32file.CreateFile(str(target), win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE, None,
                win32con.OPEN_EXISTING, 0, None)
            try:
                with rejected(OSError, "기존 대상 파일을 점유하지 못했습니다"):
                    merge.execute(overwrite_preview, overwrite=True, progress=lambda *_: None)
            finally:
                handle.Close()
            assert capture_signature(target) == before_target
            merge.execute(overwrite_preview, overwrite=True, progress=lambda *_: None)
            check_output(target, MERGED, formulas=False, broken_reference=True)

            # Re-input is allowed under a NEW destination and retains duplicates.
            checkpoint("merge prior output as input into a new destination")
            reinput = root / "결과 재입력.xlsx"
            reinput_preview = merge.preview((target, inputs[0]), reinput)
            assert reinput_preview.row_count == 8
            merge.execute(reinput_preview, overwrite=False, progress=lambda *_: None)
            check_output(reinput, MERGED + GROUPS["부산"], formulas=False, broken_reference=True)

            # Split output cannot overwrite itself when selected as a new source.
            checkpoint("Split source collision, overwrite approval and locked partial result")
            with rejected(WorkbookValidationError, "원본 파일을 덮어쓸 수 없습니다"):
                split.preview(outputs["서울"], "분류 표", "지역", "% 분할", root)
            split_preview = split.preview(source, "분류 표", "지역", "% 분할", root)
            assert set(split_preview.collisions) == set(outputs.values())
            with rejected(WorkbookValidationError, "덮어쓰기 승인"):
                split.execute(split_preview, overwrite=False, progress=lambda *_: None)
            locked = outputs["서울"]
            before_locked = capture_signature(locked)
            handle = win32file.CreateFile(str(locked), win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE, None,
                win32con.OPEN_EXISTING, 0, None)
            try:
                partial = split.execute(split_preview, overwrite=True, progress=lambda *_: None)
            finally:
                handle.Close()
            assert len(partial.failed) == 1 and partial.failed[0].label == "서울", partial
            assert set(partial.succeeded) == set(outputs.values()) - {locked}
            assert capture_signature(locked) == before_locked
            for label, path in outputs.items():
                check_output(path, GROUPS[label], formulas=True, broken_reference=True)

            # Only a disposable extra copy is changed to exercise stale preview.
            checkpoint("reject a stale Merge preview")
            changed = root / "변경 감지용 복사본.xlsx"
            shutil.copy2(inputs[0], changed)
            stale_target = root / "생성되면 안 됨.xlsx"
            stale = merge.preview((changed, inputs[1]), stale_target)
            with changed.open("ab") as stream:
                stream.write(b"synthetic stale-preview mutation")
            with rejected(WorkbookValidationError, "미리보기 이후 변경"):
                merge.execute(stale, overwrite=False, progress=lambda *_: None)
            assert not stale_target.exists()

            checkpoint("known limitation: Split rejects an actively filtered Table")
            filtered_dir = root / "필터 제한 점검"
            filtered_dir.mkdir()
            filtered_preview = split.preview(filtered, "분류 표", "지역", "% 분할", filtered_dir)
            assert filtered_preview.snapshot.row_count == 6
            try:
                split.execute(filtered_preview, overwrite=False, progress=lambda *_: None)
            except ParallelWriteAborted as exc:
                assert "필터링된 범위나 표" in str(exc), str(exc)
                assert exc.partial_result.failed, exc.partial_result
                assert set(filtered_dir.iterdir()) == set(exc.partial_result.succeeded), tuple(filtered_dir.iterdir())
                assert capture_signature(filtered) == originals[filtered]
                print("KNOWN LIMITATION: filtered Split rejected; original unchanged; temporary files removed", flush=True)
            else:
                raise AssertionError("Expected the confirmed active-filter Split rejection; review this limitation")
        finally:
            try:
                split.shutdown()
            finally:
                # Check preservation even when an earlier native assertion fails.
                actual_signatures = {path: capture_signature(path) for path in originals}
                assert actual_signatures == originals, (
                    "Original input signatures changed",
                    [str(path) for path in originals if actual_signatures[path] != originals[path]],
                )
                print("PASS: original input SHA-256, size and mtime unchanged (finally)", flush=True)
        checkpoint("temporary output cleanup")
        expected_files = {*originals, *outputs.values(), target, empty_target, reinput, changed, filtered_dir}
        assert set(root.iterdir()) == expected_files, tuple(root.iterdir())
    print("PASS: unfiltered Split -> Merge conservation/order, empty inputs, validation, overwrite, locked outputs, re-input, stale preview, unchanged originals and cleanup; active-filter rejection confirmed separately")


if __name__ == "__main__":
    main()
