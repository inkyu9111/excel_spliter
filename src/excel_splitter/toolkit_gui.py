from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .compare_service import CompareService
from .controller import AppController
from .etc_service import EtcService
from .gui import ExcelSplitterGui
from .merge_service import MergePreview, MergeService
from .naming import _unique_filename


class ExcelFileToolkitGui(ExcelSplitterGui):
    """Share the Split worker lifecycle across the toolkit's operation tabs."""

    def __init__(self, root: tk.Tk, controller: AppController,
                 merge_service: MergeService | None = None,
                 compare_service: CompareService | None = None,
                 etc_service: EtcService | None = None) -> None:
        self.merge_service = merge_service if merge_service is not None else MergeService()
        self.compare_service = compare_service if compare_service is not None else CompareService()
        self.etc_service = etc_service if etc_service is not None else EtcService()
        self.compare_tables = ((), ())
        self.merge_sources: list[Path] = []
        self.merge_preview: MergePreview | None = None
        super().__init__(root, controller)
        self.merge_output_var.trace_add("write", self._merge_output_changed)
        self.compare_output_var.trace_add("write", self._compare_output_changed)
        self.etc_output_var.trace_add("write", self._etc_output_changed)
        self._render_merge_state()
        self._render_compare_state()
        self._render_etc_state()

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
        self.merge_output_entry = ttk.Entry(output, textvariable=self.merge_output_var)
        self.merge_output_entry.grid(row=0, column=1, sticky="ew", padx=6)
        browse = ttk.Button(output, text="찾아보기", command=self._browse_merge_output)
        browse.grid(row=0, column=2)
        self._merge_widgets.extend((self.merge_output_entry, browse))
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
        self._build_compare()
        self._build_etc()

    def _build_etc(self) -> None:
        page = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(page, text="기타 (Etc)")
        page.columnconfigure(1, weight=1)
        ttk.Label(page, text="한 파일에서 선택한 시트를 정리합니다. 원본은 유지하고 새 파일로 저장합니다.",
                  wraplength=700).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self.etc_source_var = tk.StringVar()
        self.etc_sheet_var = tk.StringVar()
        self.etc_output_var = tk.StringVar()
        self.etc_remove_artifacts_var = tk.BooleanVar(value=False)
        self.etc_reset_fill_var = tk.BooleanVar(value=False)
        self.etc_exclude_table_headers_var = tk.BooleanVar(value=True)
        ttk.Label(page, text="원본 파일").grid(row=1, column=0, sticky="w")
        ttk.Entry(page, textvariable=self.etc_source_var, state="readonly").grid(row=1, column=1, sticky="ew", padx=8)
        browse = ttk.Button(page, text="찾아보기", command=self._browse_etc_source)
        browse.grid(row=1, column=2)
        ttk.Label(page, text="워크시트").grid(row=2, column=0, sticky="w", pady=8)
        self.etc_sheet_combo = ttk.Combobox(page, textvariable=self.etc_sheet_var, state="disabled")
        self.etc_sheet_combo.grid(row=2, column=1, sticky="ew", padx=8)
        self.etc_sheet_combo.bind("<<ComboboxSelected>>", lambda _: self._render_etc_state())
        self._input_widgets.extend((browse, self.etc_sheet_combo))
        self.etc_remove_artifacts_checkbox = ttk.Checkbutton(
            page, text="모든 도형·메모·댓글 삭제 (그림, 차트, 버튼 포함)",
            variable=self.etc_remove_artifacts_var, command=self._render_etc_state,
        )
        self.etc_remove_artifacts_checkbox.grid(row=3, column=0, columnspan=3, sticky="w", pady=6)
        fill_options = ttk.Frame(page)
        fill_options.grid(row=4, column=0, columnspan=3, sticky="w", pady=6)
        self.etc_reset_fill_checkbox = ttk.Checkbutton(
            fill_options, text="셀 채우기색 초기화 (표 스타일·조건부 서식 유지)",
            variable=self.etc_reset_fill_var, command=self._render_etc_state,
        )
        self.etc_reset_fill_checkbox.grid(row=0, column=0, sticky="w")
        self.etc_exclude_table_headers_checkbox = ttk.Checkbutton(
            fill_options, text="테이블 헤더 제외", variable=self.etc_exclude_table_headers_var,
            command=self._render_etc_state,
        )
        self.etc_exclude_table_headers_checkbox.grid(row=0, column=1, sticky="w", padx=(16, 0))
        self._input_widgets.extend((self.etc_remove_artifacts_checkbox, self.etc_reset_fill_checkbox,
                                    self.etc_exclude_table_headers_checkbox))
        ttk.Label(page, text="직접 지정한 채우기색만 제거하며 글자색, 테두리, 수식, 값은 유지합니다.",
                  wraplength=700).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 14))
        ttk.Label(page, text="결과 파일").grid(row=6, column=0, sticky="w")
        self.etc_output_entry = ttk.Entry(page, textvariable=self.etc_output_var)
        self.etc_output_entry.grid(row=6, column=1, sticky="ew", padx=8)
        output = ttk.Button(page, text="찾아보기", command=self._browse_etc_output)
        output.grid(row=6, column=2)
        self._input_widgets.extend((self.etc_output_entry, output))
        self.etc_progress = ttk.Progressbar(page, variable=self.progress_var, maximum=1)
        self.etc_progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(16, 6))
        self.etc_status_var = tk.StringVar(value="파일을 불러온 뒤 시트와 작업을 선택하세요.")
        ttk.Label(page, textvariable=self.etc_status_var, wraplength=700).grid(row=8, column=0, columnspan=3, sticky="w")
        self.etc_button = ttk.Button(page, text="실행 및 저장", command=self._run_etc)
        self.etc_button.grid(row=9, column=2, sticky="e", pady=12)
        self._input_widgets.append(self.etc_button)

    def _render_etc_state(self) -> None:
        ready = all((self.etc_source_var.get(), self.etc_sheet_var.get(), self.etc_output_var.get()))
        ready = ready and (self.etc_remove_artifacts_var.get() or self.etc_reset_fill_var.get())
        self.etc_exclude_table_headers_checkbox.configure(
            state="normal" if self.etc_reset_fill_var.get() and not self._busy else "disabled",
        )
        self.etc_sheet_combo.configure(state="readonly" if self.etc_sheet_combo["values"] and not self._busy else "disabled")
        self.etc_button.configure(state="normal" if ready and not self._busy else "disabled")

    def _browse_etc_source(self) -> None:
        selected = filedialog.askopenfilename(parent=self.root, title="정리할 Excel 파일 선택",
                                              filetypes=(("Excel 통합문서", "*.xlsx"),))
        if not selected:
            return
        source = Path(selected).resolve()
        self.etc_source_var.set(str(source))
        self.etc_sheet_combo.configure(values=())
        self.etc_sheet_var.set("")
        filename = _unique_filename(f"{source.stem}_정리결과", {p.name.casefold() for p in source.parent.iterdir()})
        self.etc_output_var.set(str(source.parent / filename))
        self._start_worker(lambda: ("etc_source", self.etc_service.inspect_source(source)))

    def _browse_etc_output(self) -> None:
        current = Path(self.etc_output_var.get() or "정리결과.xlsx")
        selected = filedialog.asksaveasfilename(
            parent=self.root, title="정리 결과를 저장할 새 파일", initialdir=str(current.parent),
            initialfile=current.name, defaultextension=".xlsx",
            filetypes=(("Excel 통합문서", "*.xlsx"),), confirmoverwrite=False,
        )
        if selected:
            self.etc_output_var.set(selected)

    def _etc_output_changed(self, *_: object) -> None:
        self._render_etc_state()

    def _run_etc(self) -> None:
        source, target = Path(self.etc_source_var.get()), Path(self.etc_output_var.get())
        sheet_name = self.etc_sheet_var.get()
        remove_artifacts, reset_fill = self.etc_remove_artifacts_var.get(), self.etc_reset_fill_var.get()
        exclude_table_headers = self.etc_exclude_table_headers_var.get()
        self._start_worker(lambda: ("etc_execute", self.etc_service.execute(
            source, sheet_name, target, remove_artifacts=remove_artifacts, reset_fill=reset_fill,
            exclude_table_headers=exclude_table_headers,
            progress=lambda completed, total, label: self.events.put(("progress", completed, total, label)),
        )))

    def _build_compare(self) -> None:
        page = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(page, text="비교 (Compare)")
        page.columnconfigure(1, weight=1)
        ttk.Label(page, text="기본은 같은 시트·셀 위치 비교입니다. 키 비교는 선택한 표의 행을 키로 연결합니다.",
                  wraplength=700).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Label(page, text="다른 셀은 노란색으로 표시한 새 파일에 저장합니다. 원본 파일은 그대로 유지됩니다.",
                  wraplength=700).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self.compare_reference_var = tk.StringVar()
        self.compare_comparison_var = tk.StringVar()
        self.compare_output_var = tk.StringVar()
        for row, (label, variable, action) in enumerate((
            ("기준 파일", self.compare_reference_var, lambda: self._browse_compare_input("reference")),
            ("비교대상 파일", self.compare_comparison_var, lambda: self._browse_compare_input("comparison")),
            ("결과 파일", self.compare_output_var, self._browse_compare_output),
        ), 2):
            ttk.Label(page, text=label).grid(row=row, column=0, sticky="w", pady=6)
            entry = ttk.Entry(page, textvariable=variable, state="normal" if variable is self.compare_output_var else "readonly")
            entry.grid(row=row, column=1, sticky="ew", padx=8)
            if variable is self.compare_output_var:
                self.compare_output_entry = entry
                self._input_widgets.append(entry)
            button = ttk.Button(page, text="찾아보기", command=action)
            button.grid(row=row, column=2)
            self._input_widgets.append(button)
        options = ttk.Frame(page)
        options.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        options.columnconfigure(1, weight=1)
        self.compare_by_key_var = tk.BooleanVar(value=False)
        mode = ttk.Checkbutton(options, text="키 컬럼으로 행 비교", variable=self.compare_by_key_var,
                               command=self._compare_mode_changed)
        mode.grid(row=0, column=0, sticky="w")
        self.compare_tables_button = ttk.Button(options, text="표·컬럼 불러오기", command=self._load_compare_tables)
        self.compare_tables_button.grid(row=0, column=1, sticky="e")
        self.compare_reference_table_combo = ttk.Combobox(options, state="disabled")
        self.compare_comparison_table_combo = ttk.Combobox(options, state="disabled")
        for row, (label, combo) in enumerate((
            ("기준 표", self.compare_reference_table_combo), ("비교대상 표", self.compare_comparison_table_combo),
        ), 1):
            ttk.Label(options, text=label).grid(row=row, column=0, sticky="w", pady=3)
            combo.grid(row=row, column=1, sticky="ew", padx=(8, 0))
            combo.bind("<<ComboboxSelected>>", self._refresh_compare_keys)
        ttk.Label(options, text="키 컬럼 (복수 선택)").grid(row=3, column=0, sticky="nw", pady=3)
        key_frame = ttk.Frame(options)
        key_frame.grid(row=3, column=1, sticky="ew", padx=(8, 0))
        key_frame.columnconfigure(0, weight=1)
        self.compare_key_list = tk.Listbox(key_frame, selectmode="multiple", exportselection=False, height=4)
        self.compare_key_list.grid(row=0, column=0, sticky="ew")
        scrollbar = ttk.Scrollbar(key_frame, orient="vertical", command=self.compare_key_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.compare_key_list.configure(yscrollcommand=scrollbar.set)
        self.compare_key_list.bind("<<ListboxSelect>>", lambda _: self._render_compare_state())
        self._input_widgets.extend((mode, self.compare_tables_button, self.compare_reference_table_combo,
                                    self.compare_comparison_table_combo, self.compare_key_list))
        self.compare_progress = ttk.Progressbar(page, variable=self.progress_var, maximum=1)
        self.compare_progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 6))
        self.compare_status_var = tk.StringVar(value="기준 파일과 비교대상 파일을 선택하세요.")
        ttk.Label(page, textvariable=self.compare_status_var, wraplength=700).grid(
            row=7, column=0, columnspan=3, sticky="w")
        self.compare_button = ttk.Button(page, text="비교 및 저장", command=self._compare)
        self.compare_button.grid(row=8, column=2, sticky="e", pady=8)
        self._input_widgets.append(self.compare_button)

    def _render_compare_state(self) -> None:
        ready = all(variable.get() for variable in (
            self.compare_reference_var, self.compare_comparison_var, self.compare_output_var,
        ))
        key_mode = self.compare_by_key_var.get()
        enabled = key_mode and not self._busy
        has_sources = bool(self.compare_reference_var.get() and self.compare_comparison_var.get())
        self.compare_tables_button.configure(state="normal" if enabled and has_sources else "disabled")
        for combo, tables in zip((self.compare_reference_table_combo, self.compare_comparison_table_combo), self.compare_tables):
            combo.configure(state="readonly" if enabled and tables else "disabled")
        self.compare_key_list.configure(state="normal" if enabled and self.compare_key_list.size() else "disabled")
        if key_mode:
            ready = ready and bool(self.compare_key_list.curselection())
        self.compare_button.configure(state="normal" if ready and not self._busy else "disabled")

    def _compare_mode_changed(self) -> None:
        self.compare_status_var.set("표·컬럼을 불러온 뒤 키 컬럼을 클릭하여 선택하세요." if self.compare_by_key_var.get()
                                    else "같은 이름의 시트에서 같은 위치의 셀 값을 비교합니다.")
        self._render_compare_state()

    def _invalidate_compare_tables(self) -> None:
        self.compare_tables = ((), ())
        for combo in (self.compare_reference_table_combo, self.compare_comparison_table_combo):
            combo.configure(values=())
            combo.set("")
        self._refresh_compare_keys()

    def _refresh_compare_keys(self, _event: object = None) -> None:
        self.compare_key_list.configure(state="normal")
        self.compare_key_list.delete(0, "end")
        left, right = self.compare_reference_table_combo.current(), self.compare_comparison_table_combo.current()
        if left >= 0 and right >= 0:
            for column in self.compare_tables[0][left].columns:
                self.compare_key_list.insert("end", column)
        self._render_compare_state()

    def _load_compare_tables(self) -> None:
        reference, comparison = Path(self.compare_reference_var.get()), Path(self.compare_comparison_var.get())
        self._invalidate_compare_tables()
        self._start_worker(lambda: ("compare_tables", self.compare_service.inspect_tables(reference, comparison)))

    def _browse_compare_input(self, kind: str) -> None:
        label = "기준" if kind == "reference" else "비교대상"
        selected = filedialog.askopenfilename(parent=self.root, title=f"{label} Excel 파일 선택",
                                              filetypes=(("Excel 통합문서", "*.xlsx"),))
        if not selected:
            return
        path = Path(selected).resolve()
        variable = self.compare_reference_var if kind == "reference" else self.compare_comparison_var
        variable.set(str(path))
        self._invalidate_compare_tables()
        if kind == "comparison":
            filename = _unique_filename(f"{path.stem}_비교결과", {p.name.casefold() for p in path.parent.iterdir()})
            self.compare_output_var.set(str(path.parent / filename))
        self.compare_status_var.set("파일과 결과 저장 위치를 확인하고 비교 및 저장을 누르세요.")
        self._render_compare_state()

    def _browse_compare_output(self) -> None:
        current = Path(self.compare_output_var.get() or "비교결과.xlsx")
        selected = filedialog.asksaveasfilename(
            parent=self.root, title="비교 결과를 저장할 새 파일", initialdir=str(current.parent),
            initialfile=current.name, defaultextension=".xlsx",
            filetypes=(("Excel 통합문서", "*.xlsx"),), confirmoverwrite=False,
        )
        if selected:
            self.compare_output_var.set(selected)

    def _compare_output_changed(self, *_: object) -> None:
        self._render_compare_state()

    def _compare(self) -> None:
        reference = Path(self.compare_reference_var.get())
        comparison = Path(self.compare_comparison_var.get())
        target = Path(self.compare_output_var.get())
        options = {}
        if self.compare_by_key_var.get():
            keys = tuple(self.compare_key_list.get(index) for index in self.compare_key_list.curselection())
            if not keys:
                messagebox.showerror("키 컬럼 선택", "비교할 키 컬럼을 하나 이상 선택하세요.", parent=self.root)
                return
            reference_table = self.compare_tables[0][self.compare_reference_table_combo.current()]
            comparison_table = self.compare_tables[1][self.compare_comparison_table_combo.current()]
            options = dict(key_columns=keys,
                           reference_table=(reference_table.sheet_name, reference_table.table_name),
                           comparison_table=(comparison_table.sheet_name, comparison_table.table_name))
        self._start_worker(lambda: ("compare_execute", self.compare_service.execute(
            reference, comparison, target, **options,
            progress=lambda completed, total, label: self.events.put(("progress", completed, total, label)),
        )))

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

    def _merge_output_changed(self, *_: object) -> None:
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
            self.compare_status_var.set("처리 중입니다...")
            self.etc_status_var.set("처리 중입니다...")
        self._render_merge_state()
        self._render_compare_state()
        self._render_etc_state()

    def _handle_ok(self, payload: object) -> None:
        tag, value = payload
        if tag == "etc_source":
            self.etc_sheet_combo.configure(values=value)
            self.etc_sheet_var.set(value[0] if value else "")
            self._set_busy(False)
            self.etc_status_var.set("정리할 시트와 작업을 선택하고 실행 및 저장을 누르세요.")
        elif tag == "etc_execute":
            self._set_busy(False)
            self.etc_status_var.set(f"정리 완료: {value}")
            messagebox.showinfo("시트 정리 결과", f"선택한 작업을 적용한 새 파일을 저장했습니다.\n{value}", parent=self.root)
        elif tag == "compare_tables":
            self.compare_tables = value
            for combo, tables in zip((self.compare_reference_table_combo, self.compare_comparison_table_combo), value):
                combo.configure(values=tuple(f"{table.sheet_name} / {table.table_name}" for table in tables))
                combo.current(0)
            self._refresh_compare_keys()
            self._set_busy(False)
            self.compare_status_var.set("양쪽 표를 확인하고 키 컬럼을 클릭하여 선택하세요. 수식은 계산값으로 비교합니다.")
        elif tag == "merge_preview":
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
        elif tag == "compare_execute":
            self._set_busy(False)
            summary = f"비교 완료: 다른 셀 {value.changed_cells}개를 노란색으로 표시했습니다.\n{value.target}"
            if value.missing_sheets:
                summary += "\n\n비교대상에 없어 표시하지 못한 기준 시트: " + ", ".join(value.missing_sheets)
            if value.missing_rows:
                summary += f"\n비교대상에 없는 기준 키: {value.missing_rows}개"
            if value.missing_columns:
                summary += "\n비교대상에 없는 기준 컬럼: " + ", ".join(value.missing_columns)
            self.compare_status_var.set(summary)
            messagebox.showinfo("비교 결과", summary, parent=self.root)
        else:
            super()._handle_ok(payload)

    def _handle_error(self, error: object) -> None:
        selected = self.notebook.index(self.notebook.select())
        merging = selected == 1
        if merging:
            self._invalidate_merge_preview()
        super()._handle_error(error)
        if merging:
            self.merge_status_var.set("오류가 발생했습니다. 파일과 설정을 확인한 뒤 미리보기를 다시 실행하세요.")
        elif selected == 2:
            self.compare_status_var.set("오류가 발생했습니다. 파일·표·키 컬럼과 새 결과 파일의 경로를 확인하세요.")
        elif selected == 3:
            self.etc_status_var.set("오류가 발생했습니다. 파일·시트·작업과 새 결과 파일의 경로를 확인하세요.")

    def _reset_progress(self) -> None:
        super()._reset_progress()
        self.merge_progress.configure(maximum=1)
        self.compare_progress.configure(maximum=1)
        self.etc_progress.configure(maximum=1)

    def _show_progress(self, completed: int, total: int, label: str) -> None:
        super()._show_progress(completed, total, label)
        self.merge_progress.configure(maximum=max(total, 1))
        self.merge_status_var.set(f"처리 중: {label} ({completed}/{total})")
        self.compare_progress.configure(maximum=max(total, 1))
        self.compare_status_var.set(f"처리 중: {label} ({completed}/{total})")
        self.etc_progress.configure(maximum=max(total, 1))
        self.etc_status_var.set(f"처리 중: {label} ({completed}/{total})")
