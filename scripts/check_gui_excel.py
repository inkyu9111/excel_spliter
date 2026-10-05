"""Exercise the real Tk controls, worker queue, Excel services and saved result."""

from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from time import monotonic, sleep
import tkinter as tk
from tkinter import ttk
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from check_merge_excel import _new_table
from excel_splitter.controller import AppController
from excel_splitter.excel_gateway import ExcelComGateway, _excel_session, _open_workbook
from excel_splitter.file_signature import capture_signature
from excel_splitter.split_service import SplitService
from excel_splitter.toolkit_gui import ExcelFileToolkitGui


def wait_idle(root, gui):
    deadline = monotonic() + 180
    updates = []
    while gui._busy:
        assert monotonic() < deadline, "GUI worker did not finish in 180 seconds"
        root.update()
        status = gui.status_var.get()
        if not updates or updates[-1] != status:
            updates.append(status)
        sleep(0.01)
    root.update()
    assert all(gui.notebook.tab(index, "state") == "normal" for index in range(4))
    return updates


def main():
    with TemporaryDirectory(prefix="excel-toolkit-gui-") as directory:
        folder = Path(directory)
        first, second, mismatch = (folder / name for name in ("서울.xlsx", "부산.xlsx", "다른_열.xlsx"))
        target = folder / "병합결과.xlsx"
        split_dir = folder / "분할 결과"
        split_dir.mkdir()
        expected = (("서울", 10.0), ("부산", 20.0), ("부산", 30.0))
        with _excel_session() as excel:
            for path, headers, rows in (
                (first, ("지역", "금액"), expected[:1]),
                (second, ("지역", "금액"), expected[1:]),
                (mismatch, ("다른 지역", "금액"), expected[:1]),
            ):
                with _new_table(excel, path, headers, rows):
                    pass
        sources = (first, second, mismatch)
        before = tuple(capture_signature(path) for path in sources)
        root = tk.Tk()
        root.withdraw()
        gui = ExcelFileToolkitGui(root, AppController(SplitService(ExcelComGateway())))
        selected = [str(first), str(mismatch)]
        try:
            with patch("excel_splitter.toolkit_gui.filedialog.askopenfilenames", side_effect=lambda **_: tuple(selected)), \
                 patch("excel_splitter.gui.filedialog.askopenfilename", return_value=str(target)), \
                 patch("excel_splitter.gui.filedialog.askdirectory", return_value=str(split_dir)), \
                 patch("excel_splitter.toolkit_gui.messagebox.askyesno", return_value=True) as confirm:
                gui.merge_list_buttons[0].invoke()
                gui.merge_output_var.set(str(target))
                gui.merge_button.invoke()
                assert gui._busy and gui.merge_button.instate(["disabled"])
                assert gui.notebook.tab(0, "state") == "disabled"
                wait_idle(root, gui)
                assert gui.merge_preview is None and not target.exists()
                assert "열 이름과 순서" in gui.error_detail
                assert gui.error_panel.winfo_manager() == "grid"

                selected[:] = (str(first), str(second))
                gui.merge_list_buttons[2].invoke()
                gui.merge_list_buttons[0].invoke()
                assert not gui.merge_button.instate(["disabled"])
                gui.merge_button.invoke()
                assert gui._busy
                assert gui.merge_button.instate(["disabled"])
                updates = wait_idle(root, gui)
                assert updates
                assert gui.result_paths.get("merge") == (target,), getattr(gui, "error_detail", "")
                assert gui.result_panels["merge"][0].winfo_manager() == "grid"
                assert not gui._progress_running and not gui._executing
                assert str(gui.merge_progress["mode"]) == "determinate"
                assert float(gui.merge_progress["value"]) == 1
                assert all(gui.merge_tree.set(item, "status") == "병합 완료" for item in gui.merge_tree.get_children())
                gui.result_panels["merge"][3][2].invoke()
                assert gui.notebook.index(gui.notebook.select()) == 2
                assert Path(gui.compare_comparison_var.get()) == target
                assert confirm.call_count == 0, "New merge output must run without a confirmation dialog"

                merged_signature = capture_signature(target)
                gui.notebook.select(0)
                root.deiconify()
                root.update()
                source_picker = next(widget for widget in gui.source_entry.master.winfo_children() if isinstance(widget, ttk.Button))
                source_picker.invoke()
                assert gui._busy and gui.split_button.instate(["disabled"])
                wait_idle(root, gui)
                assert gui.controller.state.source == target, getattr(gui, "error_detail", "")
                gui.sheet_combo.current(tuple(gui.sheet_combo["values"]).index("Data"))
                gui.sheet_combo.event_generate("<<ComboboxSelected>>")
                assert gui._busy
                wait_idle(root, gui)
                gui.column_combo.current(tuple(gui.column_combo["values"]).index("지역"))
                gui.column_combo.event_generate("<<ComboboxSelected>>")
                output_picker = next(widget for widget in gui.output_entry.master.winfo_children() if isinstance(widget, ttk.Button))
                output_picker.invoke()
                gui.pattern_entry.delete(0, "end")
                gui.pattern_entry.insert(0, "지역_%")
                assert gui.controller.state.column_name == "지역"
                assert gui.controller.state.output_dir == split_dir
                assert gui.controller.state.pattern == "지역_%"
                assert not gui.split_button.instate(["disabled"])
                gui.split_button.invoke()
                assert gui._busy and gui.notebook.tab(1, "state") == "disabled"
                assert wait_idle(root, gui)
                split_targets = {
                    split_dir / "지역_서울.xlsx": expected[:1],
                    split_dir / "지역_부산.xlsx": expected[1:],
                }
                assert set(gui.result_paths.get("split", ())) == set(split_targets), getattr(gui, "error_detail", "")
                assert gui.result_panels["split"][0].winfo_manager() == "grid"
                assert not gui._progress_running and not gui._executing
                assert confirm.call_count == 1, "Split should ask only for its destructive transformation warning"
                assert confirm.call_args.args[0] == "분할 확인"
        finally:
            if not gui._busy:
                gui._on_close()
            else:
                # The outer E2E process timeout handles a genuinely stuck native call.
                root.destroy()
        with _excel_session() as excel:
            for path, rows in {target: expected, **split_targets}.items():
                book = _open_workbook(excel, path, read_only=True)
                try:
                    assert book.Worksheets.Count == 1
                    actual = book.Worksheets.Item(1).ListObjects.Item(1).DataBodyRange.Value2
                    assert actual == rows, (path, actual)
                finally:
                    book.Close(SaveChanges=False)
                    book = None
        assert tuple(capture_signature(path) for path in sources) == before
        assert capture_signature(target) == merged_signature
        assert not tuple(folder.glob(".em-*"))
    print("PASS: Tk controls → validation failure → retry → native merge → compare handoff → single-click native split", flush=True)


if __name__ == "__main__":
    main()
