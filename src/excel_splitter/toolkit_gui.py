from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .controller import AppController
from .gui import ExcelSplitterGui
from .merge_service import MergePreview, MergeService


class ExcelFileToolkitGui(ExcelSplitterGui):
    """Keep the Split screen and worker lifecycle, adding a separate Merge tab."""

    def __init__(self, root: tk.Tk, controller: AppController,
                 merge_service: MergeService | None = None) -> None:
        self.merge_service = merge_service if merge_service is not None else MergeService()
        self.merge_sources: list[Path] = []
        self.merge_preview: MergePreview | None = None
        super().__init__(root, controller)
        self._render_merge_state()

    def _build(self) -> None:
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        split_page = ttk.Frame(self.notebook)
        self.notebook.add(split_page, text="분할 (Split)")
        super()._build(split_page)

        page = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(page, text="병합 (Merge)")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(2, weight=1)
        ttk.Label(page, text="같은 열 이름·순서의 Excel Table을 목록 순서대로 병합합니다.").grid(row=0, column=0, sticky="w")
        ttk.Label(page, text="첫 파일의 서식을 기준으로 사용하며, Table 수식은 현재 계산값으로 합칩니다.").grid(row=1, column=0, sticky="w", pady=(3, 10))

        table_frame = ttk.Frame(page)
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        self.merge_tree = ttk.Treeview(table_frame, columns=("path", "sheet", "rows"), show="headings", selectmode="browse")
        for key, label, width in (("path", "파일 (위에서 아래 순서)", 480), ("sheet", "워크시트", 120), ("rows", "행 수", 70)):
            self.merge_tree.heading(key, text=label)
            self.merge_tree.column(key, width=width, anchor="e" if key == "rows" else "w")
        self.merge_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.merge_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.merge_tree.configure(yscrollcommand=scrollbar.set)

        actions = ttk.Frame(page)
        actions.grid(row=3, column=0, sticky="w", pady=8)
        self._merge_widgets = []
        for column, (label, action) in enumerate((
            ("파일 추가", self._add_merge_files), ("선택 제거", self._remove_merge_file),
            ("위로", lambda: self._move_merge_file(-1)), ("아래로", lambda: self._move_merge_file(1)),
        )):
            button = ttk.Button(actions, text=label, command=action)
            button.grid(row=0, column=column, padx=(0, 6))
            self._merge_widgets.append(button)

        output = ttk.Frame(page)
        output.grid(row=4, column=0, sticky="ew", pady=4)
        output.columnconfigure(1, weight=1)
        self.merge_output_var = tk.StringVar()
        ttk.Label(output, text="결과 파일").grid(row=0, column=0)
        ttk.Entry(output, textvariable=self.merge_output_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=6)
        browse = ttk.Button(output, text="찾아보기", command=self._browse_merge_output)
        browse.grid(row=0, column=2)
        self._merge_widgets.append(browse)
        self.merge_progress = ttk.Progressbar(page, variable=self.progress_var, maximum=1)
        self.merge_progress.grid(row=5, column=0, sticky="ew", pady=(8, 3))
        self.merge_status_var = tk.StringVar(value="병합할 파일을 두 개 이상 추가하세요.")
        ttk.Label(page, textvariable=self.merge_status_var, wraplength=700).grid(row=6, column=0, sticky="w", pady=3)
        buttons = ttk.Frame(page)
        buttons.grid(row=7, column=0, sticky="e", pady=(8, 0))
        self.merge_preview_button = ttk.Button(buttons, text="미리보기", command=self._preview_merge)
        self.merge_preview_button.grid(row=0, column=0, padx=4)
        self.merge_button = ttk.Button(buttons, text="병합", command=self._merge)
        self.merge_button.grid(row=0, column=1)
        self._merge_widgets.extend((self.merge_preview_button, self.merge_button))
        self._input_widgets.extend(self._merge_widgets)

    def _render_merge_state(self) -> None:
        ready = len(self.merge_sources) >= 2 and bool(self.merge_output_var.get())
        self.merge_preview_button.configure(state="normal" if ready and not self._busy else "disabled")
        self.merge_button.configure(state="normal" if self.merge_preview is not None and not self._busy else "disabled")

    def _invalidate_merge_preview(self) -> None:
        self.merge_preview = None
        self.merge_tree.delete(*self.merge_tree.get_children())
        for index, path in enumerate(self.merge_sources):
            self.merge_tree.insert("", "end", iid=str(index), values=(str(path), "", ""))
        self.merge_status_var.set(f"파일 {len(self.merge_sources)}개 · 순서를 확인하고 미리보기를 누르세요.")
        self._render_merge_state()

    def _add_merge_files(self) -> None:
        selected = filedialog.askopenfilenames(parent=self.root, title="병합할 Excel 파일 선택", filetypes=(("Excel 통합문서", "*.xlsx"),))
        if not selected:
            return
        for name in selected:
            path = Path(name).resolve()
            if path not in self.merge_sources:
                self.merge_sources.append(path)
        if not self.merge_output_var.get():
            self.merge_output_var.set(str(self.merge_sources[0].parent / "merged.xlsx"))
        self._invalidate_merge_preview()

    def _remove_merge_file(self) -> None:
        selection = self.merge_tree.selection()
        if selection:
            del self.merge_sources[int(selection[0])]
            self._invalidate_merge_preview()

    def _move_merge_file(self, step: int) -> None:
        selection = self.merge_tree.selection()
        if not selection:
            return
        old = int(selection[0])
        new = old + step
        if 0 <= new < len(self.merge_sources):
            self.merge_sources.insert(new, self.merge_sources.pop(old))
            self._invalidate_merge_preview()
            self.merge_tree.selection_set(str(new))

    def _browse_merge_output(self) -> None:
        current = Path(self.merge_output_var.get() or "merged.xlsx")
        selected = filedialog.asksaveasfilename(parent=self.root, title="병합 결과 파일", initialdir=str(current.parent), initialfile=current.name,
                                              defaultextension=".xlsx", filetypes=(("Excel 통합문서", "*.xlsx"),), confirmoverwrite=False)
        if selected:
            self.merge_output_var.set(selected)
            self._invalidate_merge_preview()

    def _preview_merge(self) -> None:
        sources = tuple(self.merge_sources)
        target = Path(self.merge_output_var.get())
        self._invalidate_merge_preview()
        self._start_worker(lambda: ("merge_preview", self.merge_service.preview(sources, target)))

    def _merge(self) -> None:
        preview = self.merge_preview
        if preview is None:
            return
        prompt = (f"파일 {len(preview.inputs)}개, 데이터 {preview.row_count}행을 병합합니다.\n"
                  f"Table 수식은 현재 계산값으로 저장합니다.\n\n{preview.target}")
        if preview.prior_signature is not None:
            prompt += "\n\n위 기존 결과 파일을 덮어씁니다."
        if not messagebox.askyesno("병합 확인", prompt, parent=self.root):
            return
        self._start_worker(lambda: ("merge_execute", self.merge_service.execute(
            preview, preview.prior_signature is not None,
            lambda completed, total, label: self.events.put(("progress", completed, total, label)),
        )))

    def _set_busy(self, busy: bool) -> None:
        super()._set_busy(busy)
        selected = self.notebook.select()
        for tab in self.notebook.tabs():
            self.notebook.tab(tab, state="disabled" if busy and tab != selected else "normal")
        if busy:
            self._reset_progress()
            self.merge_status_var.set("처리 중입니다...")
        self._render_merge_state()

    def _handle_ok(self, payload: object) -> None:
        tag, value = payload
        if tag == "merge_preview":
            self.merge_preview = value
            for index, item in enumerate(value.inputs):
                self.merge_tree.item(str(index), values=(str(item.source), item.sheet_name, item.row_count))
            self._set_busy(False)
            self.merge_status_var.set(f"파일 {len(value.inputs)}개 · 데이터 {value.row_count}행 · 병합 준비 완료")
        elif tag == "merge_execute":
            self._invalidate_merge_preview()
            self._set_busy(False)
            self.merge_status_var.set(f"병합 완료: {value}")
            messagebox.showinfo("병합 결과", f"병합 파일을 저장했습니다.\n{value}", parent=self.root)
        else:
            super()._handle_ok(payload)

    def _handle_error(self, error: object) -> None:
        merging = self.notebook.index(self.notebook.select()) == 1
        if merging:
            self._invalidate_merge_preview()
        super()._handle_error(error)
        if merging:
            self.merge_status_var.set("오류가 발생했습니다. 파일과 설정을 확인한 뒤 미리보기를 다시 실행하세요.")

    def _reset_progress(self) -> None:
        super()._reset_progress()
        self.merge_progress.configure(maximum=1)

    def _show_progress(self, completed: int, total: int, label: str) -> None:
        super()._show_progress(completed, total, label)
        self.merge_progress.configure(maximum=max(total, 1))
        self.merge_status_var.set(f"처리 중: {label} ({completed}/{total})")
