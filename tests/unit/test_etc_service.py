from contextlib import nullcontext
from types import SimpleNamespace

import pytest

from excel_splitter.errors import SplitExecutionError, WorkbookValidationError


class _Collection:
    def __init__(self, *items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, index):
        return self.items[index - 1]


def _artifacts():
    collection = _Collection()
    for _ in range(2):
        item = SimpleNamespace()
        item.Delete = lambda item=item: collection.items.remove(item)
        collection.items.append(item)
    return collection


class _FormatConditions:
    def __init__(self, count=1):
        self.Count = count

    def Delete(self):
        self.Count = 0


def _setup(tmp_path, monkeypatch):
    from excel_splitter import etc_service as etc

    source, target = tmp_path / "source.xlsx", tmp_path / "cleaned.xlsx"
    source.write_bytes(b"source")
    events = []

    def sheet(name):
        return SimpleNamespace(Name=name, ProtectContents=False, ProtectDrawingObjects=False,
            Shapes=_artifacts(), Comments=_artifacts(), CommentsThreaded=_artifacts(),
            ListObjects=SimpleNamespace(TableStyle="TableStyleMedium2", Count=0),
            Cells=SimpleNamespace(Interior=SimpleNamespace(Pattern=1), Value2=((1, "keep"),), Formula=(("=1", "keep"),),
                Font=SimpleNamespace(Bold=True), Borders=SimpleNamespace(LineStyle=1), NumberFormat="0.00",
                FormatConditions=_FormatConditions(), ClearFormats=lambda: pytest.fail("Broad ClearFormats is forbidden")))

    selected, other = sheet("Selected"), sheet("Other")
    book = SimpleNamespace(Worksheets=_Collection(selected, other), FullName="")
    book.Close = lambda **options: events.append(("close", options))

    def save(path, **options):
        assert options["FileFormat"] == 51 and options["Password"] == ""
        from pathlib import Path
        Path(path).write_bytes(b"cleaned")
        book.FullName = path
        events.append(("save", options))

    book.SaveAs = save

    def opened(_excel, path, *, read_only):
        assert (path == source) == read_only
        if not read_only:
            assert path != target and path.read_bytes() == b"source"
        events.append(("open", read_only))
        return book

    monkeypatch.setattr(etc, "_excel_session", lambda: nullcontext(object()))
    monkeypatch.setattr(etc, "_open_workbook", opened)
    monkeypatch.setattr(etc, "_verify_xlsx_package", lambda _: None)
    return etc, source, target, selected, other, book, events


def test_inspection_reads_only_names_and_closes_original(tmp_path, monkeypatch):
    etc, source, _target, _selected, _other, _book, events = _setup(tmp_path, monkeypatch)
    assert etc.EtcService().inspect_source(source) == ("Selected", "Other")
    assert events == [("open", True), ("close", {"SaveChanges": False})]
    assert source.read_bytes() == b"source"


@pytest.mark.parametrize("remove_artifacts,reset_fill", [(True, False), (False, True), (True, True)])
def test_cleanup_options_affect_only_selected_sheet_and_keep_other_cell_styles(tmp_path, monkeypatch, remove_artifacts, reset_fill):
    etc, source, target, selected, other, _book, events = _setup(tmp_path, monkeypatch)
    progress = []
    output = etc.EtcService().execute(source, "Selected", target,
        remove_artifacts=remove_artifacts, reset_fill=reset_fill, progress=lambda *args: progress.append(args))
    assert output == target and target.read_bytes() == b"cleaned" and source.read_bytes() == b"source"
    assert set(tmp_path.iterdir()) == {source, target}
    assert (selected.Shapes.Count, selected.Comments.Count, selected.CommentsThreaded.Count) == ((0, 0, 0) if remove_artifacts else (2, 2, 2))
    assert selected.Cells.Interior.Pattern == (-4142 if reset_fill else 1)
    assert selected.ListObjects.TableStyle == "TableStyleMedium2"
    assert selected.Cells.FormatConditions.Count == 1
    assert selected.Cells.Value2 == ((1, "keep"),) and selected.Cells.Formula == (("=1", "keep"),)
    assert selected.Cells.Font.Bold and selected.Cells.Borders.LineStyle == 1 and selected.Cells.NumberFormat == "0.00"
    assert (other.Shapes.Count, other.Comments.Count, other.CommentsThreaded.Count, other.Cells.Interior.Pattern) == (2, 2, 2, 1)
    assert events[0] == ("open", False) and events[-1] == ("close", {"SaveChanges": False})
    assert progress[-1] == (1, 1, "Selected")


def test_conditional_format_removal_is_independent_and_scoped_to_selected_sheet(tmp_path, monkeypatch):
    etc, source, target, selected, other, _book, _events = _setup(tmp_path, monkeypatch)

    etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=False,
                             remove_conditional_formats=True, progress=lambda *_: None)

    assert selected.Cells.FormatConditions.Count == 0
    assert other.Cells.FormatConditions.Count == 1
    assert selected.Cells.Interior.Pattern == other.Cells.Interior.Pattern == 1
    assert selected.Shapes.Count == other.Shapes.Count == 2


def test_conditional_format_removal_also_clears_table_headers_when_fill_excludes_them(tmp_path, monkeypatch):
    etc, source, target, selected, _other, _book, _events = _setup(tmp_path, monkeypatch)
    fills = _track_fill_cells(selected, 4, 5)
    selected.Cells.FormatConditions = _FormatConditions()
    selected.ListObjects = _Collection(_table(2, 2, 3))

    etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=True,
                             remove_conditional_formats=True, exclude_table_headers=True, progress=lambda *_: None)

    assert {cell for cell, pattern in fills.items() if pattern == 1} == {(2, 2), (2, 3), (2, 4)}
    assert selected.Cells.FormatConditions.Count == 0


def test_conditional_format_deletion_failure_keeps_original_and_never_publishes(tmp_path, monkeypatch):
    etc, source, target, selected, _other, _book, events = _setup(tmp_path, monkeypatch)
    selected.Cells.FormatConditions.Delete = lambda: (_ for _ in ()).throw(RuntimeError("delete failed"))

    with pytest.raises(SplitExecutionError):
        etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=False,
                                 remove_conditional_formats=True, progress=lambda *_: None)

    assert source.read_bytes() == b"source"
    assert not target.exists()
    assert {path.name for path in tmp_path.iterdir()} == {"source.xlsx"}
    assert events == [("open", False), ("close", {"SaveChanges": False})]


@pytest.mark.parametrize("invalid", ["noop", "sheet", "protected_cells", "protected_objects", "source_target", "existing", "hardlink", "symlink", "dangling_symlink"])
def test_rejection_preserves_original_and_existing_output(tmp_path, monkeypatch, invalid):
    etc, source, target, selected, _other, _book, _events = _setup(tmp_path, monkeypatch)
    sheet_name = "Missing" if invalid == "sheet" else "Selected"
    if invalid == "protected_cells":
        selected.ProtectContents = True
    elif invalid == "protected_objects":
        selected.ProtectDrawingObjects = True
    elif invalid == "source_target":
        target = source
    elif invalid == "existing":
        target.write_bytes(b"prior")
    elif invalid == "hardlink":
        target.hardlink_to(source)
    elif invalid == "symlink":
        target.symlink_to(source)
    elif invalid == "dangling_symlink":
        target.symlink_to(tmp_path / "missing.xlsx")
    before = set(tmp_path.iterdir())
    with pytest.raises(WorkbookValidationError):
        etc.EtcService().execute(source, sheet_name, target, remove_artifacts=invalid != "noop", reset_fill=invalid != "noop", progress=lambda *_: None)
    assert set(tmp_path.iterdir()) == before
    assert source.read_bytes() == b"source"
    if invalid == "existing":
        assert target.read_bytes() == b"prior"
    assert selected.Shapes.Count == 2 and selected.Cells.Interior.Pattern == 1


@pytest.mark.parametrize("failure", ["save", "package", "source_changed", "target_created", "copy"])
def test_failure_never_publishes_partial_output_and_closes_copy(tmp_path, monkeypatch, failure):
    etc, source, target, _selected, _other, book, events = _setup(tmp_path, monkeypatch)

    def broken(*_args, **_options):
        raise RuntimeError("injected failure")

    if failure == "save":
        book.SaveAs = broken
    elif failure == "package":
        monkeypatch.setattr(etc, "_verify_xlsx_package", broken)
    elif failure == "copy":
        monkeypatch.setattr(etc.shutil, "copy2", lambda _source, temp: temp.write_bytes(b"bad copy"))

    def progress(*_):
        if failure == "source_changed":
            source.write_bytes(b"external change")
        elif failure == "target_created":
            target.write_bytes(b"new owner")

    expected = SplitExecutionError if failure == "save" else (RuntimeError, WorkbookValidationError, OSError)
    with pytest.raises(expected) as captured:
        etc.EtcService().execute(source, "Selected", target, remove_artifacts=True, reset_fill=True, progress=progress)
    if failure == "save":
        assert "결과 저장" in str(captured.value)
        assert isinstance(captured.value.__cause__, RuntimeError)
    assert source.read_bytes() == (b"external change" if failure == "source_changed" else b"source")
    assert set(tmp_path.iterdir()) == ({source, target} if failure == "target_created" else {source})
    if failure == "target_created":
        assert target.read_bytes() == b"new owner"
    if failure != "copy":
        assert events[-1] == ("close", {"SaveChanges": False})


def test_fill_only_does_not_reject_object_protection(tmp_path, monkeypatch):
    etc, source, target, selected, _other, _book, _events = _setup(tmp_path, monkeypatch)
    selected.ProtectDrawingObjects = True
    etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=True, progress=lambda *_: None)
    assert target.exists() and selected.Shapes.Count == 2 and selected.Cells.Interior.Pattern == -4142


def _track_fill_cells(sheet, rows, columns):
    fills = {(row, column): 1 for row in range(1, rows + 1) for column in range(1, columns + 1)}

    class Interior:
        def __init__(self, cells):
            self.cells = cells

        @property
        def Pattern(self):
            return None

        @Pattern.setter
        def Pattern(self, value):
            for cell in self.cells:
                fills[cell] = value

    cells = SimpleNamespace(Interior=Interior(tuple(fills)))
    cells.Item = lambda row, column: SimpleNamespace(Row=row, Column=column)

    def cell_range(first, last):
        return SimpleNamespace(Interior=Interior(tuple(
            (row, column)
            for row in range(first.Row, last.Row + 1)
            for column in range(first.Column, last.Column + 1)
        )))

    sheet.Rows = SimpleNamespace(Count=rows)
    sheet.Columns = SimpleNamespace(Count=columns)
    sheet.Cells = cells
    sheet.Range = cell_range
    return fills


def _table(row, column, width, *, show_headers=True, header_range=True):
    header = SimpleNamespace(Row=row, Column=column, Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=width))
    return SimpleNamespace(ShowHeaders=show_headers, HeaderRowRange=header if header_range else None)


def test_reset_fill_preserves_only_visible_table_header_cells_across_multiple_layouts(tmp_path, monkeypatch):
    etc, source, target, selected, _other, _book, _events = _setup(tmp_path, monkeypatch)
    fills = _track_fill_cells(selected, 6, 9)
    selected.ListObjects = _Collection(
        _table(2, 2, 3), _table(2, 6, 2), _table(4, 1, 2),
    )

    etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=True,
                             progress=lambda *_: None)

    headers = {(2, column) for column in range(2, 5)} | {(2, column) for column in range(6, 8)} | {(4, column) for column in range(1, 3)}
    assert {cell for cell, pattern in fills.items() if pattern == 1} == headers
    assert {cell for cell, pattern in fills.items() if pattern == -4142} == set(fills) - headers


def test_reset_fill_can_clear_visible_table_headers_when_exclusion_is_unchecked(tmp_path, monkeypatch):
    etc, source, target, selected, _other, _book, _events = _setup(tmp_path, monkeypatch)
    fills = _track_fill_cells(selected, 4, 5)
    selected.ListObjects = _Collection(_table(2, 2, 3))

    etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=True,
                             exclude_table_headers=False, progress=lambda *_: None)

    assert set(fills.values()) == {-4142}


@pytest.mark.parametrize("tables", [
    _Collection(),
    _Collection(_table(2, 2, 3, show_headers=False)),
    _Collection(_table(2, 2, 3, header_range=False)),
])
def test_reset_fill_clears_entire_sheet_when_a_table_header_is_unavailable(tmp_path, monkeypatch, tables):
    etc, source, target, selected, _other, _book, _events = _setup(tmp_path, monkeypatch)
    fills = _track_fill_cells(selected, 4, 5)
    selected.ListObjects = tables

    etc.EtcService().execute(source, "Selected", target, remove_artifacts=False, reset_fill=True,
                             exclude_table_headers=True, progress=lambda *_: None)

    assert set(fills.values()) == {-4142}
