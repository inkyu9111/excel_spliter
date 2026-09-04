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

    def inspect(_workbook, source, signature):
        return merge.MergeInput(
            source, signature, "Data", "Table1", ("Name", "Amount"), rows[source], 1, False
        )

    monkeypatch.setattr(merge, "_inspect_source", inspect)
    monkeypatch.setattr(merge, "_open_workbook", lambda *_args, **_kwargs: _workbook()[0])
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

    def incompatible(workbook, source, signature):
        item = inspect(workbook, source, signature)
        return item if source == sources[0] else replace(item, columns=("Amount", "Name"))

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


def test_execute_reports_copy_and_merge_stages_before_completion(tmp_path, monkeypatch):
    merge, sources, target, _rows = _setup(tmp_path, monkeypatch)
    preview = merge.MergeService().preview(sources, target)
    monkeypatch.setattr(merge, "_write_merged", lambda temp, _preview, _progress: temp.write_bytes(b"merged"))
    progress = []

    merge.MergeService().execute(preview, overwrite=False, progress=lambda *event: progress.append(event))

    assert progress == [
        (0, 0, "원본 확인 중"),
        (0, 0, "파일 복사 중"),
        (0, 0, "값 병합 중"),
        (1, 1, "완료"),
    ]


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

    def __init__(self, table, start=0, rows=None, header=False, columns=None):
        self.table, self.start, self.header = table, start, header
        self.rows = len(table.data) if rows is None else rows
        self.Row = 3 if header else 4 + start
        self.Column = 2
        self.columns = len(table.headers) if columns is None else columns
        self.Rows = SimpleNamespace(Count=self.rows)
        self.Columns = SimpleNamespace(Count=self.columns)
        self.MergeCells = False
        self.Hyperlinks = _Collection()

    @property
    def Resize(self):
        raise AssertionError("Range.Resize is an optional-argument COM property, not a method")

    def Cells(self, row, column):
        assert column == 1
        return _Range(self.table, self.start + row - 1, 1, columns=1)

    @property
    def Value2(self):
        if self.header:
            self.table.header_reads += 1
            return (self.table.headers,) if self.columns > 1 else self.table.headers[0]
        return tuple(tuple(cell[0] for cell in row) for row in self.table.data[self.start:self.start + self.rows])

    @property
    def Formula(self):
        return tuple(tuple(cell[1] for cell in row) for row in self.table.data[self.start:self.start + self.rows])

    def Copy(self):
        _Range.clipboard = deepcopy(self.table.data[self.start:self.start + self.rows])

    def PasteSpecial(self, *, Paste):
        assert Paste in (12, -4163, -4122), "formulas must never cross workbook boundaries"
        assert self.rows == len(_Range.clipboard)
        assert all(len(row) == self.columns for row in _Range.clipboard)
        self.table.pastes.append((self.Row, self.Column, self.rows, self.columns, Paste))
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

    def __init__(self, data, totals, headers=("Name", "Amount")):
        self.data, self.ShowTotals = deepcopy(data), totals
        self.headers, self.header_reads = headers, 0
        self.HeaderRowRange = _Range(self, rows=1, header=True)
        self.totals = "=SUBTOTAL(109,[Amount])" if totals else None
        self.resizes, self.pastes = [], []

    @property
    def ListRows(self):
        return SimpleNamespace(Count=len(self.data))

    @property
    def DataBodyRange(self):
        return _Range(self) if self.data else None

    @property
    def Range(self):
        return SimpleNamespace(Row=3, Column=2, Rows=SimpleNamespace(Count=1 + len(self.data) + int(self.ShowTotals)), Columns=SimpleNamespace(Count=len(self.headers)))

    def Resize(self, area):
        assert (area.Row, area.Column, area.columns) == (3, 2, len(self.headers))
        self.resizes.append((area.Row, area.Column, area.rows, area.columns))
        count = area.rows - 1 - int(self.ShowTotals)
        self.data = self.data[:count] + [[[None, None, "General"] for _ in self.headers] for _ in range(max(0, count - len(self.data)))]


@pytest.mark.parametrize("totals", [False, True])
@pytest.mark.parametrize("kind", ["normal", "empty", "empty_first", "one_column", "no_growth"])
def test_native_writer_keeps_values_order_duplicates_formats_and_first_totals(tmp_path, monkeypatch, kind, totals):
    from excel_splitter import merge_service as merge
    sources = (tmp_path / "first.xlsx", tmp_path / "second.xlsx")
    temp = tmp_path / "temporary.xlsx"
    data = [
        [[["repeat", None, "General"], [4, "=2+2", "0.00"]]] * 2,
        [[["repeat", None, "General"], [4, "=1+3", "0.000"]],
         [["literal", None, "General"], ["=unsafe", None, "@"]],
         [["error", None, "General"], [-2146826281, "=1/0", "General"]]],
    ]
    headers = ("Name", "Amount")
    if kind == "empty":
        data = [[], []]
    elif kind == "empty_first":
        data[0] = []
    elif kind == "one_column":
        headers = ("Name",)
        data = [[[["first", None, "@"]]], [[["second", None, "General"]]]]
    elif kind == "no_growth":
        data[1] = []
    books, tables = {}, {}
    for path, values, show_totals in ((temp, data[0], totals), (sources[0], data[0], totals), (sources[1], data[1], False)):
        book, sheet, _old_table = _workbook()
        table = _Table(values, show_totals, headers)
        sheet.ListObjects = _Collection(table)
        def rectangle(first, last, table=table):
            assert first[1] == 2
            return _Range(table, start=first[0] - 4, rows=last[0] - first[0] + 1,
                          header=first[0] == 3, columns=last[1] - first[1] + 1)
        sheet.Range = rectangle
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
    inputs = tuple(merge.MergeInput(path, object(), "Data", "DataTable", headers, len(rows), 3, totals if index == 0 else False) for index, (path, rows) in enumerate(zip(sources, data)))
    preview = merge.MergePreview(inputs, tmp_path / "output.xlsx", None, sum(len(rows) for rows in data))
    progress = []
    merge._write_merged(temp, preview, lambda *event: progress.append(event))
    output = tables[temp]
    assert [[cell[0] for cell in row] for row in output.data] == [[cell[0] for cell in row] for row in data[0] + data[1]]
    assert all(cell[1] is None for row in output.data for cell in row)
    assert [[cell[2] for cell in row] for row in output.data] == [[cell[2] for cell in row] for row in data[0] + data[1]]
    assert output.totals == ("=SUBTOTAL(109,[Amount])" if totals else None)
    expected_shape = {
        "normal": (6, 2, [(4, 2, 2, 2), (6, 2, 3, 2)]),
        "empty": (None, 2, []),
        "empty_first": (4, 2, [(4, 2, 3, 2)]),
        "one_column": (3, 1, [(4, 2, 1, 1), (5, 2, 1, 1)]),
        "no_growth": (None, 2, [(4, 2, 2, 2)]),
    }[kind]
    resize_rows, width, destinations = expected_shape
    assert output.resizes == ([] if resize_rows is None else [(3, 2, resize_rows + int(totals), width)])
    assert [area[:4] for area in output.pastes if area[4] == 12] == destinations
    assert all(tables[path].header_reads == 1 for path in sources)
    assert books[temp].saved and all(book.closed for book in books.values())
    assert books[temp].save_options == dict(FileFormat=51, Password="", WriteResPassword="", ReadOnlyRecommended=False, AddToMru=False)
    assert all(not books[path].saved for path in sources)
    assert progress[:2] == [(1, 2, "first.xlsx"), (2, 2, "second.xlsx")]
    assert progress[2:] == [(0, 0, "결과 저장 중"), (0, 0, "저장 결과 확인 중")]


class _Header:
    Row, Column = 3, 2

    def __init__(self, values=(("Name", "Amount"),)):
        self.values, self.reads = values, 0

    @property
    def Value2(self):
        self.reads += 1
        return self.values


def _workbook():
    table = SimpleNamespace(
        Name="DataTable", SourceType=1, ShowHeaders=True, ShowTotals=False,
        ListColumns=_Collection(SimpleNamespace(Name="Name"), SimpleNamespace(Name="Amount")),
        ListRows=SimpleNamespace(Count=2), HeaderRowRange=_Header(),
        Range=SimpleNamespace(Row=3, Column=2, Rows=SimpleNamespace(Count=3), Columns=SimpleNamespace(Count=2)),
    )
    sheet = SimpleNamespace(
        Name="Data", Visible=-1, ProtectContents=False, ListObjects=_Collection(table),
        UsedRange=SimpleNamespace(Row=1, Rows=SimpleNamespace(Count=50)),
        Comments=_Collection(), CommentsThreaded=_Collection(), FilterMode=False,
        content={}, ranges=[],
    )
    sheet.Cells = lambda row, column: (row, column)
    sheet.Range = lambda first, last: _sheet_range(sheet, first, last)
    book = SimpleNamespace(ProtectStructure=False, Sheets=_Collection(sheet), Worksheets=_Collection(sheet), closed=False, open_count=0, close_count=0)
    def close(*, SaveChanges):
        assert SaveChanges is False
        book.closed = True
        book.close_count += 1
    book.Close = close
    return book, sheet, table


def _sheet_range(sheet, first, last):
    """COM range double with values and artifacts at real worksheet coordinates."""
    sheet.ranges.append((first, last))
    rows = []
    for row in range(first[0], last[0] + 1):
        cells = []
        for column in range(first[1], last[1] + 1):
            properties = dict(Value2=None, Formula=None, MergeCells=False,
                              Hyperlinks=_Collection(), Comment=None, CommentThreaded=None, Style="Normal")
            properties.update(sheet.content.get((row, column), {}))
            cells.append(SimpleNamespace(**properties))
        rows.append(cells)
    cells = [cell for row in rows for cell in row]
    return SimpleNamespace(
        Value2=tuple(tuple(cell.Value2 for cell in row) for row in rows),
        Formula=tuple(tuple(cell.Formula for cell in row) for row in rows),
        MergeCells=cells[0].MergeCells if all(cell.MergeCells == cells[0].MergeCells for cell in cells) else None,
        Hyperlinks=SimpleNamespace(Count=sum(cell.Hyperlinks.Count for cell in cells)),
        Style=cells[0].Style if all(cell.Style == cells[0].Style for cell in cells) else None,
        Cells=_Collection(*cells),
    )


def _preview_workbooks(tmp_path, monkeypatch, *, totals=False, row_counts=(2, 3)):
    from excel_splitter import merge_service as merge

    sources = (tmp_path / "first.xlsx", tmp_path / "second.xlsx")
    books = {}
    for index, source in enumerate(sources):
        source.write_bytes(f"source {index}".encode())
        book, sheet, table = _workbook()
        table.ListRows.Count = row_counts[index]
        table.ShowTotals = totals if index == 0 else False
        table.Range.Rows.Count = 1 + row_counts[index] + int(table.ShowTotals)
        books[source] = book
    monkeypatch.setattr(merge, "_excel_session", lambda: nullcontext(object()))
    def opened(_excel, source, *, read_only):
        assert read_only is True
        books[source].open_count += 1
        books[source].closed = False
        return books[source]
    monkeypatch.setattr(merge, "_open_workbook", opened)
    return merge, sources, tmp_path / "merged.xlsx", books


@pytest.mark.parametrize("headers", [(("Name", "Amount"),), "Name"])
def test_preview_opens_each_source_once_and_reads_headers_in_bulk(tmp_path, monkeypatch, headers):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch)
    first = books[sources[0]]
    for book in books.values():
        table = book.Worksheets.Item(1).ListObjects.Item(1)
        table.HeaderRowRange = _Header(headers)
        table.ListColumns.Item = lambda *_: pytest.fail("header names must be read in bulk")
    validate = merge._validate_merge_expansion
    def expansion(sheet, table, rows):
        assert not first.closed
        assert books[sources[1]].closed
        validate(sheet, table, rows)
    monkeypatch.setattr(merge, "_validate_merge_expansion", expansion)
    preview = merge.MergeService().preview(sources, target)
    assert preview.inputs[0].columns == (("Name", "Amount") if isinstance(headers, tuple) else ("Name",))
    assert [book.open_count for book in books.values()] == [1, 1]
    assert [book.close_count for book in books.values()] == [1, 1]
    assert all(book.Worksheets.Item(1).ListObjects.Item(1).HeaderRowRange.reads == 1 for book in books.values())


def test_preview_accepts_same_workbook_name_in_different_directories(tmp_path, monkeypatch):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch)
    renamed = []
    for index, source in enumerate(sources):
        path = tmp_path / str(index) / "same.xlsx"
        path.parent.mkdir()
        path.write_bytes(source.read_bytes())
        books[path] = books.pop(source)
        renamed.append(path)
    sources = tuple(renamed)
    opened = merge._open_workbook
    def open_excel(excel, source, **kwargs):
        for path, book in books.items():
            if book.open_count > book.close_count and path.name.casefold() == source.name.casefold():
                raise RuntimeError("Excel cannot open two workbooks with the same name")
        return opened(excel, source, **kwargs)
    monkeypatch.setattr(merge, "_open_workbook", open_excel)
    preview = merge.MergeService().preview(sources, target)
    assert tuple(item.source for item in preview.inputs) == sources
    assert tuple(item.row_count for item in preview.inputs) == (2, 3)
    assert preview.row_count == 5
    assert all(book.open_count == book.close_count == 1 for book in books.values())


@pytest.mark.parametrize("failure", ["open", "inspect", "schema", "expansion"])
def test_preview_closes_each_opened_book_on_error(tmp_path, monkeypatch, failure):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch)
    first, second = (books[source] for source in sources)
    if failure == "open":
        opened = merge._open_workbook
        def fail_open(excel, source, **kwargs):
            if sum(book.open_count for book in books.values()) == 1:
                raise RuntimeError("open failed")
            return opened(excel, source, **kwargs)
        monkeypatch.setattr(merge, "_open_workbook", fail_open)
    elif failure == "inspect":
        second.Worksheets.Item(1).ProtectContents = True
    elif failure == "schema":
        table = second.Worksheets.Item(1).ListObjects.Item(1)
        table.HeaderRowRange.values = (("Amount", "Name"),)
        table.ListColumns = _Collection(SimpleNamespace(Name="Amount"), SimpleNamespace(Name="Name"))
    else:
        first.Worksheets.Item(1).content[6, 2] = {"Value2": "keep"}
    with pytest.raises((RuntimeError, WorkbookValidationError)):
        merge.MergeService().preview(sources, target)
    assert sum(book.open_count for book in books.values()) >= 1
    assert all(book.close_count == book.open_count for book in books.values())
    assert all(book.closed for book in books.values() if book.open_count)
    assert [source.read_bytes() for source in sources] == [b"source 0", b"source 1"]


@pytest.mark.parametrize("style", ["SplitResidual", "표준", "mixed"])
def test_preview_accepts_styled_blank_split_remnants(tmp_path, monkeypatch, style):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch)
    sheet = books[sources[0]].Worksheets.Item(1)
    for row in range(6, 51):
        for column in (2, 3):
            sheet.content[row, column] = {"Style": "Normal" if style == "mixed" and row % 2 else style}
    preview = merge.MergeService().preview(sources, target)
    assert preview.row_count == 5
    assert sheet.ranges == [((6, 2), (8, 3))]
    assert all(book.closed for book in books.values())


@pytest.mark.parametrize("totals, first_row, last_row", [(False, 6, 8), (True, 7, 9)])
def test_preview_checks_only_new_table_rows_and_columns(tmp_path, monkeypatch, totals, first_row, last_row):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch, totals=totals)
    sheet = books[sources[0]].Worksheets.Item(1)
    sheet.content = {
        (first_row - 1, 2): {"Value2": "existing table or totals"},
        (last_row + 1, 2): {"Value2": "below final table"},
        (first_row, 1): {"Value2": "left"},
        (first_row, 4): {"Value2": "right"},
        (50, 2): {"Comment": object(), "CommentThreaded": object()},
    }
    sheet.Comments = sheet.CommentsThreaded = _Collection(SimpleNamespace(Parent=SimpleNamespace(Row=50, Column=2)))
    preview = merge.MergeService().preview(sources, target)
    assert preview.row_count == 5
    assert sheet.ranges == [((first_row, 2), (last_row, 3))]


@pytest.mark.parametrize("content", [
    {"Value2": "keep"}, {"Value2": 0}, {"Value2": False},
    {"Value2": "", "Formula": '=IF(TRUE,"",1)'},
    {"Hyperlinks": _Collection(object())}, {"MergeCells": True},
    {"Comment": object()}, {"CommentThreaded": object()},
])
@pytest.mark.parametrize("coordinate", [(7, 2), (9, 3)])
def test_preview_rejects_content_at_both_expansion_boundaries(tmp_path, monkeypatch, content, coordinate):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch, totals=True)
    sheet = books[sources[0]].Worksheets.Item(1)
    sheet.content[coordinate] = content
    if "Comment" in content:
        sheet.Comments = _Collection(SimpleNamespace(Parent=SimpleNamespace(Row=coordinate[0], Column=coordinate[1])))
    if "CommentThreaded" in content:
        sheet.CommentsThreaded = _Collection(SimpleNamespace(Parent=SimpleNamespace(Row=coordinate[0], Column=coordinate[1])))
    with pytest.raises(WorkbookValidationError, match="병합.*확장"):
        merge.MergeService().preview(sources, target)
    assert sheet.ranges == [((7, 2), (9, 3))]
    assert all(book.closed for book in books.values())


@pytest.mark.parametrize("row_counts", [(0, 0), (2, 0)])
def test_preview_without_table_growth_ignores_below_table_content(tmp_path, monkeypatch, row_counts):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch, row_counts=row_counts)
    sheet = books[sources[0]].Worksheets.Item(1)
    sheet.content[6, 2] = {"Value2": "must remain"}
    assert merge.MergeService().preview(sources, target).row_count == row_counts[0]
    assert sheet.ranges == []


def test_execute_rechecks_expansion_before_resize_and_keeps_output(tmp_path, monkeypatch):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch)
    target.write_bytes(b"prior")
    service = merge.MergeService()
    preview = service.preview(sources, target)
    book = books[sources[0]]
    sheet = book.Worksheets.Item(1)
    sheet.content[6, 2] = {"Formula": '=IF(TRUE,"",1)'}
    sheet.ListObjects.Item(1).Resize = lambda *_: pytest.fail("collision reached Resize")
    def opened(_excel, source, *, read_only):
        assert source.name.startswith(".em-") and read_only is False
        return book
    monkeypatch.setattr(merge, "_open_workbook", opened)
    with pytest.raises(WorkbookValidationError, match="병합.*확장"):
        service.execute(preview, overwrite=True, progress=lambda *_: None)
    assert target.read_bytes() == b"prior"
    assert [source.read_bytes() for source in sources] == [b"source 0", b"source 1"]
    assert set(tmp_path.iterdir()) == {*sources, target}
    assert book.closed


@pytest.mark.parametrize("invalid", ["sheets", "charts", "tables", "protected", "external", "headers"])
def test_inspection_rejects_unsupported_workbook_and_closes_it(tmp_path, monkeypatch, invalid):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch)
    book = books[sources[0]]
    sheet = book.Worksheets.Item(1)
    table = sheet.ListObjects.Item(1)
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
    with pytest.raises(WorkbookValidationError):
        merge.MergeService().preview(sources, target)
    assert book.closed


def test_inspection_accepts_empty_tables_and_reads_sources_read_only(tmp_path, monkeypatch):
    merge, sources, target, books = _preview_workbooks(tmp_path, monkeypatch, row_counts=(0, 0))
    item = merge.MergeService().preview(sources, target).inputs[0]
    assert (item.row_count, item.header_row, item.columns) == (0, 3, ("Name", "Amount"))
    assert all(book.closed for book in books.values())


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
