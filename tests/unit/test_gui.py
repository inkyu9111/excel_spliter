from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_splitter.gui import DELETED_SHEETS_WARNING, ExcelSplitterGui
from excel_splitter.models import FileSignature, Preview, WorkbookSnapshot


@pytest.mark.parametrize("has_artifacts", (False, True))
def test_split_confirmation_warns_only_when_artifacts_will_be_deleted(
    has_artifacts: bool, monkeypatch: pytest.MonkeyPatch
) -> None:
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "abc"),
        "Data",
        "Orders",
        "Team",
        0,
        (),
        has_removable_artifacts=has_artifacts,
    )
    preview = Preview(snapshot, (), ())
    gui = ExcelSplitterGui.__new__(ExcelSplitterGui)
    gui.controller = SimpleNamespace(state=SimpleNamespace(preview=preview))
    gui.root = object()
    prompts: list[str] = []
    monkeypatch.setattr(
        "excel_splitter.gui.messagebox.askyesno",
        lambda _title, prompt, **_options: prompts.append(prompt) or False,
    )

    gui._split()

    expected = DELETED_SHEETS_WARNING
    if has_artifacts:
        expected += "\n\n메모 및 도형은 삭제됩니다"
    assert prompts == [expected]
