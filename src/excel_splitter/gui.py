from __future__ import annotations

import logging
import os
import queue
import threading
import tkinter as tk
from collections.abc import Callable
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .controller import AppController, UiState
from .errors import ExcelSplitterError, WorkbookValidationError
from .models import Preview, SplitResult
from .parallel_writer import ParallelWriteAborted

DELETED_SHEETS_WARNING = (
    "선택한 시트를 제외한 다른 시트를 삭제합니다. 삭제되는 시트에 의존하는 "
    "수식, 이름, 차트, 유효성 검사 및 연결이 손상될 수 있습니다."
)
UNEXPECTED_ERROR_MESSAGE = "처리 중 예상하지 못한 오류가 발생했습니다."


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("excel_splitter")
    if logger.handlers:
        return logger
    local_app_data = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    log_dir = local_app_data / "ExcelFileToolkit" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(
        log_dir / "excel-file-toolkit.log",
        maxBytes=1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    return logger


def _log_path(logger: logging.Logger) -> Path:
    for handler in getattr(logger, "handlers", ()):
        if path := getattr(handler, "baseFilename", None):
            return Path(path)
    local_app_data = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    return local_app_data / "ExcelFileToolkit" / "logs" / "excel-file-toolkit.log"


class ExcelSplitterGui:
    def __init__(self, root: tk.Tk, controller: AppController) -> None:
        self.root = root
        self.controller = controller
        self.logger = configure_logging()
        self.events: queue.Queue[tuple[object, ...]] = queue.Queue()
        self._busy = False
        self._input_widgets: list[tk.Widget] = []

        self.source_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.column_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.pattern_var = tk.StringVar(value=controller.state.pattern)
        self.progress_var = tk.DoubleVar(value=0)
        self.status_var = tk.StringVar(value="원본 파일을 선택하세요.")

        self._build()
        self._render_state(controller.state)
        self.output_var.trace_add("write", self._output_changed)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(75, self.poll_queue)

    def _build(self, parent: tk.Misc | None = None) -> None:
        self.root.title("Excel File Toolkit")
        self.root.minsize(760, 560)
        parent = parent if parent is not None else self.root
        frame = ttk.Frame(parent, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

        ttk.Label(frame, text="원본 파일").grid(row=0, column=0, sticky="w", pady=3)
        self.source_entry = ttk.Entry(frame, textvariable=self.source_var, state="readonly")
        self.source_entry.grid(row=0, column=1, sticky="ew", padx=6)
        source_button = ttk.Button(frame, text="찾아보기", command=self._browse_source)
        source_button.grid(row=0, column=2)

        ttk.Label(frame, text="워크시트").grid(row=1, column=0, sticky="w", pady=3)
        self.sheet_combo = ttk.Combobox(
            frame, textvariable=self.sheet_var, state="readonly"
        )
        self.sheet_combo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._select_sheet)

        ttk.Label(frame, text="분류 컬럼").grid(row=2, column=0, sticky="w", pady=3)
        self.column_combo = ttk.Combobox(
            frame, textvariable=self.column_var, state="readonly"
        )
        self.column_combo.grid(row=2, column=1, columnspan=2, sticky="ew", padx=6)
        self.column_combo.bind("<<ComboboxSelected>>", self._select_column)

        ttk.Label(frame, text="파일명 패턴").grid(row=3, column=0, sticky="w", pady=3)
        self.pattern_entry = ttk.Entry(frame, textvariable=self.pattern_var)
        self.pattern_entry.grid(row=3, column=1, sticky="ew", padx=6)
        self.pattern_entry.bind("<KeyRelease>", self._pattern_changed)

        ttk.Label(frame, text="출력 폴더").grid(row=4, column=0, sticky="w", pady=3)
        self.output_entry = ttk.Entry(frame, textvariable=self.output_var)
        self.output_entry.grid(row=4, column=1, sticky="ew", padx=6)
        output_button = ttk.Button(frame, text="찾아보기", command=self._browse_output)
        output_button.grid(row=4, column=2)

        self.preview_tree = ttk.Treeview(
            frame,
            columns=("label", "count", "filename"),
            show="headings",
            height=12,
        )
        self.preview_tree.heading("label", text="분류")
        self.preview_tree.heading("count", text="행 수")
        self.preview_tree.heading("filename", text="출력 파일")
        self.preview_tree.column("label", width=180)
        self.preview_tree.column("count", width=80, anchor="e")
        self.preview_tree.column("filename", width=390)
        self.preview_tree.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(10, 6))

        self.progress = ttk.Progressbar(frame, variable=self.progress_var, maximum=1)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=3)
        ttk.Label(frame, textvariable=self.status_var).grid(
            row=7, column=0, columnspan=3, sticky="w", pady=3
        )

        buttons = ttk.Frame(frame)
        buttons.grid(row=8, column=0, columnspan=3, sticky="e", pady=(8, 0))
        self.preview_button = ttk.Button(buttons, text="미리보기", command=self._preview)
        self.preview_button.grid(row=0, column=0, padx=4)
        self.split_button = ttk.Button(buttons, text="분할", command=self._split)
        self.split_button.grid(row=0, column=1)

        self._input_widgets.extend(
            [
                source_button,
                self.sheet_combo,
                self.column_combo,
                self.pattern_entry,
                self.output_entry,
                output_button,
                self.preview_button,
                self.split_button,
            ]
        )

    def _browse_source(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="원본 Excel 파일 선택",
            filetypes=(("Excel 통합문서", "*.xlsx"),),
        )
        if not selected:
            return
        self.source_var.set(selected)
        self._clear_preview()
        self._start_worker(
            lambda: ("source", self.controller.select_source(Path(selected)))
        )

    def _select_sheet(self, _event: object = None) -> None:
        name = self.sheet_var.get()
        self._clear_preview()
        self._start_worker(lambda: ("sheet", self.controller.select_sheet(name)))

    def _select_column(self, _event: object = None) -> None:
        self.controller.select_column(self.column_var.get())
        self._clear_preview()
        self._render_state(self.controller.state)

    def _pattern_changed(self, _event: object = None) -> None:
        self.controller.set_pattern(self.pattern_var.get())
        self._clear_preview()
        self._render_state(self.controller.state)

    def _browse_output(self) -> None:
        selected = filedialog.askdirectory(parent=self.root, title="출력 폴더 선택")
        if not selected:
            return
        self.controller.select_output_dir(Path(selected))
        self._clear_preview()
        self._render_state(self.controller.state)

    def _output_changed(self, *_: object) -> None:
        raw = self.output_var.get()
        output_dir = Path(raw) if raw else None
        if output_dir == self.controller.state.output_dir:
            return
        self.controller.select_output_dir(output_dir)
        self._clear_preview()
        self._render_state(self.controller.state)

    def _preview(self) -> None:
        self.controller.set_pattern(self.pattern_var.get())
        self._clear_preview()
        output_dir = self.controller.state.output_dir
        if output_dir is not None:
            try:
                if not output_dir.exists():
                    prompt = f"폴더가 존재하지 않습니다. 만드시겠습니까?\n\n{output_dir}"
                    if not messagebox.askyesno("폴더 생성", prompt, parent=self.root):
                        return
                    output_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                self._handle_error(
                    WorkbookValidationError(
                        f"출력 폴더를 확인하거나 만들 수 없습니다: {output_dir}\n{exc}"
                    )
                )
                return
        self._start_worker(lambda: ("preview", self.controller.create_preview()))

    def _split(self) -> None:
        preview = self.controller.state.preview
        if preview is None:
            return
        prompt = DELETED_SHEETS_WARNING
        if preview.snapshot.has_removable_artifacts:
            prompt += "\n\n메모 및 도형은 삭제됩니다"
        if preview.collisions:
            collision_list = "\n".join(str(path) for path in preview.collisions)
            prompt += "\n\n다음 기존 파일을 덮어씁니다:\n" + collision_list
        if not messagebox.askyesno("분할 확인", prompt, parent=self.root):
            return
        overwrite = bool(preview.collisions)

        def execute() -> tuple[str, SplitResult]:
            result = self.controller.execute(
                overwrite,
                lambda completed, total, label: self.events.put(
                    ("progress", completed, total, label)
                ),
            )
            return ("execute", result)

        self._start_worker(execute)

    def _start_worker(self, action: Callable[[], object]) -> None:
        self._set_busy(True)

        def run() -> None:
            try:
                self.events.put(("ok", action()))
            except Exception as exc:
                self.events.put(("error", exc))

        threading.Thread(target=run, daemon=True).start()

    def poll_queue(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    self._show_progress(int(event[1]), int(event[2]), str(event[3]))
                elif kind == "ok":
                    self._handle_ok(event[1])
                elif kind == "error":
                    self._handle_error(event[1])
        except queue.Empty:
            pass
        self.root.after(75, self.poll_queue)

    def _handle_ok(self, payload: object) -> None:
        self._set_busy(False)
        tag, value = payload
        if tag in {"source", "sheet"}:
            self._clear_preview()
            self._render_state(value)
            self.status_var.set("다음 항목을 선택하세요.")
        elif tag == "preview":
            self._render_preview(value)
            self._render_state(self.controller.state)
            self.status_var.set("미리보기가 준비되었습니다.")
        elif tag == "execute":
            self._show_summary(value)
            self._render_state(self.controller.state)
    def _handle_error(self, error: object) -> None:
        self._set_busy(False)
        exc = error if isinstance(error, Exception) else Exception(str(error))
        self.logger.error(
            "GUI 작업 실패",
            exc_info=(type(exc), exc, exc.__traceback__),
        )
        if isinstance(exc, ParallelWriteAborted):
            message = (
                f"{exc}\n\n완료 {len(exc.partial_result.succeeded)}개, "
                f"실패 {len(exc.partial_result.failed)}개, "
                f"시작하지 못함 {len(exc.unstarted)}개"
            )
        else:
            message = str(exc) if isinstance(exc, ExcelSplitterError) else UNEXPECTED_ERROR_MESSAGE
        message += f"\n\n자세한 내용: {_log_path(self.logger)}"
        self._reset_progress()
        messagebox.showerror("오류", message, parent=self.root)
        self._render_state(self.controller.state)
        self.status_var.set("오류가 발생했습니다.")

    def _reset_progress(self) -> None:
        self.progress.configure(maximum=1)
        self.progress_var.set(0)

    def _show_progress(self, completed: int, total: int, label: str) -> None:
        self.progress.configure(maximum=max(total, 1))
        self.progress_var.set(completed)
        shown = label if label else "∅ (빈 셀)"
        self.status_var.set(f"처리 중: {shown} ({completed}/{total})")

    def _render_preview(self, preview: Preview) -> None:
        self._clear_preview()
        counts = {group.key: group.count for group in preview.snapshot.groups}
        for target in preview.targets:
            label = target.label if target.label else "∅ (빈 셀)"
            self.preview_tree.insert(
                "", "end", values=(label, counts[target.key], target.path.name)
            )

    def _show_summary(self, result: SplitResult) -> None:
        lines = ["성공 파일:"]
        lines.extend(str(path) for path in result.succeeded)
        if not result.succeeded:
            lines.append("없음")
        lines.append("\n실패:")
        lines.extend(f"{failure.label or '∅ (빈 셀)'}: {failure.message}" for failure in result.failed)
        if not result.failed:
            lines.append("없음")
        messagebox.showinfo("분할 결과", "\n".join(lines), parent=self.root)
        self.status_var.set("분할 작업이 끝났습니다.")

    def _clear_preview(self) -> None:
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.split_button.configure(state="disabled")

    def _render_state(self, state: UiState) -> None:
        self.source_var.set(str(state.source) if state.source else "")
        raw_output = self.output_var.get()
        if (Path(raw_output) if raw_output else None) != state.output_dir:
            self.output_var.set(str(state.output_dir) if state.output_dir else "")
        self.sheet_combo.configure(values=state.sheets)
        self.column_combo.configure(values=state.columns)
        self.sheet_var.set(state.sheet_name or "")
        self.column_var.set(state.column_name or "")
        self.pattern_var.set(state.pattern)
        if self._busy:
            return
        self.sheet_combo.configure(state="readonly" if state.sheets else "disabled")
        self.column_combo.configure(state="readonly" if state.columns else "disabled")
        self.pattern_entry.configure(state="normal" if state.source else "disabled")
        ready = all(
            (state.source, state.sheet_name, state.column_name, state.output_dir)
        )
        self.preview_button.configure(state="normal" if ready else "disabled")
        self.split_button.configure(state="normal" if state.preview else "disabled")

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if busy:
            for widget in self._input_widgets:
                widget.configure(state="disabled")
            self.root.protocol("WM_DELETE_WINDOW", lambda: None)
            self.status_var.set("처리 중입니다...")
        else:
            for widget in self._input_widgets:
                widget.configure(state="normal")
            self.root.protocol("WM_DELETE_WINDOW", self._on_close)
            self._render_state(self.controller.state)

    def _on_close(self) -> None:
        if self._busy:
            return
        try:
            self.controller.shutdown()
        except Exception:
            self.logger.exception("source session 종료 실패")
        finally:
            self.root.destroy()
