from contextlib import nullcontext
from copy import deepcopy
from dataclasses import replace
from pathlib import Path
import stat
from types import SimpleNamespace
from zipfile import ZipFile

import pytest

from excel_splitter.errors import WorkbookValidationError


class _Collection:
    def __init__(self, *items):
        self.items = items
        self.Count = len(items)

    def Item(self, index):
        return self.items[index - 1]


def _setup(tmp_path, monkeypatch):
    from excel_splitter import merge_service as merge

    sources = (tmp_path / "first.xlsx", tmp_path / "second.xlsx")
    for index, source in enumerate(sources):
        source.write_bytes(f"source {index}".encode())
    monkeypatch.setattr(merge, "_excel_session", lambda: nullcontext(object()))
    rows = {sources[0]: 2, sources[1]: 3}

    def inspect(_excel, source, signature, *, template):
        return merge.MergeInput(
            source, signature, "Data", "Table1", ("Name", "Amount"), rows[source], 1, False
        )

    monkeypatch.setattr(merge, "_inspect_source", inspect)
    return merge, sources, tmp_path / "merged.xlsx", rows


def test_preview_reports_each_source_and_total_in_input_order(tmp_path, monkeypatch):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    preview = merge.MergeService().preview(sources, target)
    assert tuple(item.source for item in preview.inputs) == sources
    assert tuple(item.row_count for item in preview.inputs) == (2, 3)
    assert preview.row_count == 5
    assert preview.prior_signature is None


@pytest.mark.parametrize("kind", ["single", "duplicate", "hardlink", "same_target", "hardlink_target", "extension", "directory", "reserved", "brackets"])
def test_preview_rejects_invalid_paths_before_excel(tmp_path, monkeypatch, kind):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    if kind == "single":
        sources = sources[:1]
    elif kind == "duplicate":
        sources = (sources[0], sources[0])
    elif kind == "hardlink":
        alias = tmp_path / "alias.xlsx"
        alias.hardlink_to(sources[0])
        sources = (sources[0], alias)
    elif kind == "same_target":
        target = sources[0]
    elif kind == "hardlink_target":
        target.hardlink_to(sources[1])
    elif kind == "extension":
        target = tmp_path / "merged.xlsm"
    elif kind == "directory":
        target.mkdir()
    elif kind == "reserved":
        target = tmp_path / "CON.xlsx"
    elif kind == "brackets":
        target = tmp_path / "[merged].xlsx"
    monkeypatch.setattr(merge, "_excel_session", lambda: pytest.fail("invalid paths reached Excel"))
    with pytest.raises(WorkbookValidationError):
        merge.MergeService().preview(sources, target)


def test_preview_rejects_schema_order_mismatch(tmp_path, monkeypatch):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    inspect = merge._inspect_source

    def incompatible(excel, source, signature, *, template):
        item = inspect(excel, source, signature, template=template)
        return item if template else replace(item, columns=("Amount", "Name"))

    monkeypatch.setattr(merge, "_inspect_source", incompatible)
    with pytest.raises(WorkbookValidationError, match="열 이름과 순서"):
        merge.MergeService().preview(sources, target)


def test_preview_rejects_excel_row_limit_including_header_and_totals(tmp_path, monkeypatch):
    merge, sources, target, rows = _setup(tmp_path, monkeypatch)
    rows[sources[0]] = 1_048_570
    rows[sources[1]] = 5
    inspect = merge._inspect_source
    monkeypatch.setattr(merge, "_inspect_source", lambda *args, **kwargs: replace(inspect(*args, **kwargs), has_totals=True))
    with pytest.raises(WorkbookValidationError, match="행 한도"):
        merge.MergeService().preview(sources, target)


@pytest.mark.parametrize("changed", ["source", "target", "new_target", "target_deleted", "overwrite"])
def test_execute_rejects_stale_or_unapproved_preview_before_writing(tmp_path, monkeypatch, changed):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    if changed in ("target", "target_deleted", "overwrite"):
        target.write_bytes(b"prior")
    service = merge.MergeService()
    preview = service.preview(sources, target)
    if changed == "source":
        sources[1].write_bytes(b"changed source")
    elif changed in ("target", "new_target"):
        target.write_bytes(b"changed target")
    elif changed == "target_deleted":
        target.unlink()
    monkeypatch.setattr(merge, "_write_merged", lambda *_: pytest.fail("stale preview reached writer"))
    with pytest.raises(WorkbookValidationError):
        service.execute(preview, overwrite=changed != "overwrite", progress=lambda *_: None)


def test_execute_publishes_only_complete_copy_and_preserves_sources(tmp_path, monkeypatch):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    target.write_bytes(b"prior")
    service = merge.MergeService()
    preview = service.preview(sources, target)

    def write(temp, actual_preview, progress):
        assert temp.parent == target.parent
        assert temp.read_bytes() == b"source 0"
        assert target.read_bytes() == b"prior"
        assert actual_preview == preview
        temp.write_bytes(b"complete merge")

    monkeypatch.setattr(merge, "_write_merged", write)
    result = service.execute(preview, overwrite=True, progress=lambda *_: None)
    assert result == target
    assert target.read_bytes() == b"complete merge"
    assert [path.read_bytes() for path in sources] == [b"source 0", b"source 1"]
    assert set(tmp_path.iterdir()) == {*sources, target}


def test_execute_can_write_temp_when_first_source_is_read_only(tmp_path, monkeypatch):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    sources[0].chmod(stat.S_IREAD)
    try:
        service = merge.MergeService()
        preview = service.preview(sources, target)
        def write(temp, _preview, _progress):
            temp.write_bytes(b"merged")
        monkeypatch.setattr(merge, "_write_merged", write)
        service.execute(preview, overwrite=False, progress=lambda *_: None)
        assert target.read_bytes() == b"merged"
        assert not sources[0].stat().st_mode & stat.S_IWRITE
    finally:
        sources[0].chmod(stat.S_IREAD | stat.S_IWRITE)


class _Range:
    """Small external COM double: native clipboard operations retain cell types."""
    clipboard = None

    def __init__(self, table, start=0, rows=None, header=False):
        self.table, self.start, self.header = table, start, header
        self.rows = len(table.data) if rows is None else rows
        self.Row = 3

    def Resize(self, rows, _columns):
        return _Range(self.table, self.start, rows, self.header)

    def Cells(self, row, _column):
        return _Range(self.table, self.start + row - 1, 1)

    def Copy(self):
        _Range.clipboard = deepcopy(self.table.data[self.start:self.start + self.rows])

    def PasteSpecial(self, *, Paste):
        assert Paste in (12, -4163, -4122), "formulas must never cross workbook boundaries"
        for offset, row in enumerate(_Range.clipboard):
            for column, cell in enumerate(row):
                dest = self.table.data[self.start + offset][column]
                if Paste in (12, -4163):
                    dest[0], dest[1] = cell[0], None
                if Paste in (12, -4122):
                    dest[2] = cell[2]


class _Table:
    Name, SourceType, ShowHeaders = "DataTable", 1, True
    ListColumns = _Collection(SimpleNamespace(Name="Name"), SimpleNamespace(Name="Amount"))

    def __init__(self, data, totals):
        self.data, self.ShowTotals = deepcopy(data), totals
        self.HeaderRowRange = _Range(self, header=True)
        self.totals = "=SUBTOTAL(109,[Amount])" if totals else None

    @property
    def ListRows(self):
        return SimpleNamespace(Count=len(self.data))

    @property
    def DataBodyRange(self):
        return _Range(self) if self.data else None

    def Resize(self, area):
        count = area.rows - 1 - int(self.ShowTotals)
        self.data = self.data[:count] + [[[None, None, "General"] for _ in range(2)] for _ in range(max(0, count - len(self.data)))]


@pytest.mark.parametrize("empty", [False, True])
def test_native_writer_keeps_values_order_duplicates_formats_and_first_totals(tmp_path, monkeypatch, empty):
    from excel_splitter import merge_service as merge
    sources = (tmp_path / "first.xlsx", tmp_path / "second.xlsx")
    temp = tmp_path / "temporary.xlsx"
    data = [
        [[["repeat", None, "General"], [4, "=2+2", "0.00"]]] * 2,
        [[["repeat", None, "General"], [4, "=1+3", "0.000"]],
         [["literal", None, "General"], ["=unsafe", None, "@"]],
         [["error", None, "General"], [-2146826281, "=1/0", "General"]]],
    ] if not empty else [[], []]
    books, tables = {}, {}
    for path, values, totals in ((temp, data[0], True), (sources[0], data[0], True), (sources[1], data[1], False)):
        book, sheet, _old_table = _workbook()
        table = _Table(values, totals)
        sheet.ListObjects = _Collection(table)
        sheet.FilterMode = False
        sheet.Calculate = lambda: None
        book.saved = False
        def save(book=book):
            book.saved = True
        book.Save = save
        def save_as(filename, *, book=book, **options):
            book.saved = True
            book.save_options = options
            book.FullName = filename
            with ZipFile(filename, "w") as package:
                for part in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"):
                    package.writestr(part, b"saved by native Excel")
        book.SaveAs = save_as
        books[path], tables[path] = book, table
    excel = SimpleNamespace(CutCopyMode=False)
    monkeypatch.setattr(merge, "_excel_session", lambda: nullcontext(excel))
    def opened(_excel, path, *, read_only):
        assert read_only is (path != temp)
        return books[path]
    monkeypatch.setattr(merge, "_open_workbook", opened)
    inputs = tuple(merge.MergeInput(path, object(), "Data", "DataTable", ("Name", "Amount"), len(rows), 3, index == 0) for index, (path, rows) in enumerate(zip(sources, data)))
    preview = merge.MergePreview(inputs, tmp_path / "output.xlsx", None, sum(len(rows) for rows in data))
    progress = []
    merge._write_merged(temp, preview, lambda *event: progress.append(event))
    output = tables[temp]
    assert [[cell[0] for cell in row] for row in output.data] == ([] if empty else [["repeat", 4], ["repeat", 4], ["repeat", 4], ["literal", "=unsafe"], ["error", -2146826281]])
    assert all(cell[1] is None for row in output.data for cell in row)
    assert [row[1][2] for row in output.data] == ([] if empty else ["0.00", "0.00", "0.000", "@", "General"])
    assert output.totals == "=SUBTOTAL(109,[Amount])"
    assert books[temp].saved and all(book.closed for book in books.values())
    assert books[temp].save_options == dict(FileFormat=51, Password="", WriteResPassword="", ReadOnlyRecommended=False, AddToMru=False)
    assert all(not books[path].saved for path in sources)
    assert progress[-1] == (2, 2, "second.xlsx")


def _workbook():
    table = SimpleNamespace(
        Name="DataTable", SourceType=1, ShowHeaders=True, ShowTotals=False,
        ListColumns=_Collection(SimpleNamespace(Name="Name"), SimpleNamespace(Name="Amount")),
        ListRows=SimpleNamespace(Count=2), HeaderRowRange=SimpleNamespace(Row=3),
    )
    sheet = SimpleNamespace(Name="Data", Visible=-1, ProtectContents=False, ListObjects=_Collection(table))
    book = SimpleNamespace(ProtectStructure=False, Sheets=_Collection(sheet), Worksheets=_Collection(sheet), closed=False)
    def close(*, SaveChanges):
        assert SaveChanges is False
        book.closed = True
    book.Close = close
    return book, sheet, table


@pytest.mark.parametrize("invalid", ["sheets", "charts", "tables", "protected", "external", "headers"])
def test_inspection_rejects_unsupported_workbook_and_closes_it(tmp_path, monkeypatch, invalid):
    from excel_splitter import merge_service as merge
    book, sheet, table = _workbook()
    if invalid == "sheets":
        book.Worksheets = _Collection(sheet, sheet)
    elif invalid == "charts":
        book.Sheets = _Collection(sheet, object())
    elif invalid == "tables":
        sheet.ListObjects = _Collection(table, table)
    elif invalid == "protected":
        sheet.ProtectContents = True
    elif invalid == "external":
        table.SourceType = 0
    else:
        table.ShowHeaders = False
    monkeypatch.setattr(merge, "_open_workbook", lambda *_args, **_kwargs: book)
    monkeypatch.setattr(merge, "_validate_below_table", lambda *_: None)
    with pytest.raises(WorkbookValidationError):
        merge._inspect_source(object(), tmp_path / "input.xlsx", object(), template=True)
    assert book.closed


def test_inspection_accepts_empty_tables_and_reads_sources_read_only(tmp_path, monkeypatch):
    from excel_splitter import merge_service as merge
    book, _sheet, table = _workbook()
    table.ListRows.Count = 0
    def opened(_excel, _source, *, read_only):
        assert read_only is True
        return book
    monkeypatch.setattr(merge, "_open_workbook", opened)
    monkeypatch.setattr(merge, "_validate_below_table", lambda *_: None)
    item = merge._inspect_source(object(), tmp_path / "input.xlsx", object(), template=True)
    assert (item.row_count, item.header_row, item.columns) == (0, 3, ("Name", "Amount"))
    assert book.closed


@pytest.mark.parametrize("failure", ["writer", "source_changed", "target_changed"])
def test_execute_failure_keeps_existing_output_and_cleans_temp(tmp_path, monkeypatch, failure):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    target.write_bytes(b"prior")
    service = merge.MergeService()
    preview = service.preview(sources, target)

    def write(temp, _preview, _progress):
        temp.write_bytes(b"partial")
        if failure == "writer":
            raise RuntimeError("write failed")
        if failure == "source_changed":
            sources[1].write_bytes(b"changed source")
        else:
            target.write_bytes(b"new owner")

    monkeypatch.setattr(merge, "_write_merged", write)
    with pytest.raises((RuntimeError, WorkbookValidationError)):
        service.execute(preview, overwrite=True, progress=lambda *_: None)
    assert target.read_bytes() == (b"new owner" if failure == "target_changed" else b"prior")
    assert set(tmp_path.iterdir()) == {*sources, target}
