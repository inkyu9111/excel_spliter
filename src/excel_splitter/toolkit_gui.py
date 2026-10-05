from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from .compare_service import CompareService
from .controller import AppController
from .etc_service import EtcService
from .gui import ExcelSplitterGui
from .merge_service import MergePreview, MergeService
from .naming import _unique_filename
from .ui_helpers import suggest_output_path, validate_output_path
from .gui import _log_path


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
        self._merge_after_preview = False
        self.source_name_labels = {}
        self._progress_variables = []
        super().__init__(root, controller)
        self.merge_output_var.trace_add("write", self._merge_output_changed)
        self.compare_output_var.trace_add("write", self._compare_output_changed)
        self.etc_output_var.trace_add("write", self._etc_output_changed)
        self._render_merge_state()
        self._render_compare_state()
        self._render_etc_state()

    def _build(self) -> None:
        self.root.title("Excel File Toolkit · 파일 작업")
        self.root.iconbitmap(str(Path(__file__).with_name("assets") / "app.ico"))
        self.root.geometry("900x680")
        self.root.minsize(760, 540)
        self.root.option_add("*Font", ("맑은 고딕", 10))
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self.root.configure(background="#f4f5f7")
        style.configure(".", font=("맑은 고딕", 10), background="#ffffff", foreground="#17202b")
        style.configure("Shell.TFrame", background="#f4f5f7")
        style.configure("Shell.TLabel", background="#f4f5f7", foreground="#64748b")
        style.configure("TFrame", background="#ffffff")
        style.configure("TFrame", bordercolor="#d5dde3", lightcolor="#d5dde3", darkcolor="#d5dde3")
        style.configure("TLabel", background="#ffffff")
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor="#d5dde3", lightcolor="#d5dde3", darkcolor="#d5dde3", padding=3)
        style.map("TEntry", fieldbackground=[("readonly", "#f7f9fa")])
        style.configure("TCombobox", fieldbackground="#ffffff", background="#f7f9fa", bordercolor="#d5dde3", lightcolor="#d5dde3", darkcolor="#d5dde3", padding=3)
        style.map("TCombobox", fieldbackground=[("readonly", "#ffffff")])
        style.configure("Horizontal.TProgressbar", troughcolor="#eef2f1", background="#18754c", bordercolor="#eef2f1", lightcolor="#18754c", darkcolor="#18754c")
        style.configure("TButton", padding=(10, 5), background="#ffffff")
        style.configure("Primary.TButton", background="#1266ad", foreground="white", padding=(16, 7))
        style.map("Primary.TButton", background=[("disabled", "#e4eaf0"), ("active", "#0c508a")],
                  foreground=[("disabled", "#718279")])
        style.configure("TNotebook", background="#f4f5f7", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 8))
        style.map("TNotebook.Tab", padding=[("selected", (18, 8)), ("!selected", (18, 8))],
                  background=[("selected", "#ffffff"), ("!selected", "#e9edf2")], foreground=[("selected", "#1266ad")])
        style.configure("Treeview", rowheight=30, fieldbackground="#ffffff", bordercolor="#d5dde3")
        style.map("Treeview", background=[("selected", "#1266ad")], foreground=[("selected", "white")])
        style.configure("Treeview.Heading", padding=5, background="#f3f6f8")
        style.configure("Section.TLabel", font=("맑은 고딕", 11, "bold"))
        style.configure("Note.TLabel", foreground="#64788a")
        header = ttk.Frame(self.root, padding=(12, 10), style="Shell.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Excel 파일 도구", font=("맑은 고딕", 13, "bold"), style="Shell.TLabel",
                  foreground="#17202b").pack(side="left")
        ttk.Button(header, text="로그 보기", command=self._show_log).pack(side="right")
        ttk.Label(header, text="원본 보존 · 검증 후 결과 저장", style="Shell.TLabel").pack(side="right", padx=12)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        split_page = self._page("파일 분할")
        self._build_split_page(split_page)
        self._build_merge()
        self._build_compare()
        self._build_etc()
        self._build_error_panel()
        self._render_compare_state()
        self.notebook.select(1)

    def _page(self, title: str):
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text=title)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        canvas = tk.Canvas(outer, highlightthickness=0, background="#f4f5f7")
        canvas.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        bar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=bar.set)
        page = ttk.Frame(canvas, padding=(0, 10, 0, 4), style="Shell.TFrame")
        window = canvas.create_window(0, 0, window=page, anchor="nw")
        page.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        # Scroll only the active page; list/table widgets retain their own wheel behavior.
        def wheel(event):
            if self.notebook.select() == str(outer) and not isinstance(event.widget, (tk.Listbox, ttk.Treeview)):
                if page.winfo_height() > canvas.winfo_height():
                    canvas.yview_scroll(-int(event.delta / 120), "units")
        self.root.bind("<MouseWheel>", wheel, add="+")
        page.columnconfigure(0, weight=1)
        return page

    def _section(self, page, row: int, title: str):
        section = ttk.Frame(page, padding=(12, 8), relief="solid", borderwidth=1)
        section.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        section.columnconfigure(0, weight=1)
        ttk.Label(section, text=title, style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(2, 5))
        body = ttk.Frame(section)
        body.grid(row=1, column=0, sticky="ew")
        body.columnconfigure(0, weight=1)
        return body

    def _wrapped_label(self, parent, **options):
        label = ttk.Label(parent, width=1, justify="left", **options)
        label.bind("<Configure>", lambda event: label.configure(wraplength=max(1, event.width)))
        return label

    def _path_row(self, parent, variable, action, label="", readonly=False):
        row = ttk.Frame(parent)
        row.grid(sticky="ew", pady=4)
        row.columnconfigure(0, weight=1)
        caption = ttk.Frame(row)
        caption.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 3))
        caption.columnconfigure(1, weight=1)
        if label:
            ttk.Label(caption, text=label).grid(row=0, column=0, sticky="w", padx=(0, 8))
        if readonly:
            name = self._wrapped_label(caption, font=("맑은 고딕", 10, "bold"), anchor="e")
            name.grid(row=0, column=1, sticky="ew")
            self.source_name_labels[str(variable)] = name
            variable.trace_add("write", lambda *_: name.configure(text=Path(variable.get()).name if variable.get() else ""))
        entry = ttk.Entry(row, textvariable=variable, state="readonly" if readonly else "normal")
        entry.grid(row=1, column=0, sticky="ew", padx=(0, 8), ipady=2)
        button = ttk.Button(row, text="찾아보기", command=action)
        button.grid(row=1, column=2)
        full_path = self._wrapped_label(row, textvariable=variable, style="Note.TLabel")
        path_font = tkfont.Font(root=self.root, font=entry.cget("font"))

        def show_full_path(*_):
            if variable.get() and path_font.measure(variable.get()) > entry.winfo_width() - 12:
                full_path.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(2, 0))
            else:
                full_path.grid_remove()

        variable.trace_add("write", show_full_path)
        entry.bind("<Configure>", show_full_path)
        self._input_widgets.append(button)
        if not readonly:
            self._input_widgets.append(entry)
        return entry, row

    def _output_section(self, page, kind, variable, action):
        body = self._section(page, 2, "3   저장 위치")
        entry, row = self._path_row(body, variable, action, "결과 파일")
        recommend = ttk.Button(row, text="새 파일명 추천", command=lambda: self._recommend_output(kind))
        recommend.grid(row=1, column=1, padx=(0, 8))
        self._input_widgets.append(recommend)
        note = tk.StringVar()
        setattr(self, kind + "_path_note", note)
        self._wrapped_label(body, textvariable=note, style="Note.TLabel").grid(sticky="ew", pady=(4, 0))
        setattr(self, kind + "_output_entry", entry)

    def _work_footer(self, page, kind, text, action):
        # Keep progress and the next action visible even when settings scroll.
        footer = ttk.Frame(page.master.master, padding=(12, 10))
        footer.grid(row=1, column=0, columnspan=2, sticky="ew")
        footer.columnconfigure(0, weight=1)
        status = tk.StringVar(value="파일과 작업 설정을 확인하세요.")
        reason = tk.StringVar()
        setattr(self, kind + "_status_var", status)
        setattr(self, kind + "_reason_var", reason)
        variable = tk.DoubleVar(value=0)
        self._progress_variables.append(variable)
        progress = ttk.Progressbar(footer, variable=variable, maximum=1)
        progress.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        setattr(self, kind + "_progress", progress)
        self._wrapped_label(footer, textvariable=status, style="Note.TLabel").grid(row=1, column=0, sticky="ew")
        self._wrapped_label(footer, textvariable=reason, style="Note.TLabel").grid(row=2, column=0, sticky="ew")
        button = ttk.Button(footer, text=text, command=action, style="Primary.TButton")
        button.grid(row=1, column=2, rowspan=2)
        self._input_widgets.append(button)
        setattr(self, kind + "_button", button)
        self._build_result_panel(page, kind)

    def _build_result_panel(self, page, kind):
        frame = ttk.Frame(page, padding=12, relief="solid", borderwidth=1)
        frame.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        frame.columnconfigure(0, weight=1)
        label = self._wrapped_label(frame, font=("맑은 고딕", 10, "bold"))
        label.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        table = ttk.Frame(frame)
        table.grid(row=1, column=0, sticky="ew")
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=("kind", "location", "column", "reference", "comparison"), show="headings", height=6)
        for name, heading, width in zip(tree["columns"], ("구분", "시트 · 키 / 위치", "열", "기준 파일", "대상 파일"), (90, 220, 130, 190, 190)):
            tree.heading(name, text=heading)
            tree.column(name, width=width, minwidth=60)
        tree.grid(row=0, column=0, sticky="ew")
        bar = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        hbar = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hbar.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=bar.set, xscrollcommand=hbar.set)
        tree.bind("<Double-1>", lambda _: self._show_result_detail(kind))
        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, sticky="e", pady=(10, 0))
        buttons = []
        for i, (label_text, action) in enumerate((
            ("선택 결과 열기" if kind == "split" else "결과 파일 열기", lambda: self._open_result(kind)),
            ("결과 폴더 열기", lambda: self._open_result(kind, folder=True)),
        )):
            button = ttk.Button(actions, text=label_text, command=action)
            button.grid(row=0, column=i, padx=(6, 0))
            buttons.append(button)
        if kind == "merge":
            button = ttk.Button(actions, text="이 결과 비교하기", command=self._use_merge_for_compare)
            button.grid(row=0, column=2, padx=(6, 0))
            buttons.append(button)
        ttk.Button(actions, text="설정으로 돌아가기", command=lambda: page.master.yview_moveto(0)).grid(
            row=0, column=2 if kind in ("split", "compare") else 3, padx=(6, 0))
        if kind in ("split", "compare"):
            ttk.Button(actions, text="선택 내역 상세 보기", command=lambda: self._show_result_detail(kind)).grid(
                row=0, column=3, padx=(6, 0))
        self._input_widgets.extend(buttons)
        self.result_panels[kind] = (frame, label, tree, buttons)
        frame.grid_remove()

    def _build_split_page(self, page):
        body = self._section(page, 0, "1   파일")
        self.source_entry, _ = self._path_row(body, self.source_var, self._browse_source, "원본 파일", readonly=True)
        body = self._section(page, 1, "2   작업 설정")
        options = ttk.Frame(body)
        options.grid(sticky="ew")
        options.columnconfigure((0, 1), weight=1)
        ttk.Label(options, text="워크시트").grid(row=0, column=0, sticky="w")
        ttk.Label(options, text="분류 컬럼").grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.sheet_combo = ttk.Combobox(options, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=1, column=0, sticky="ew", pady=5)
        self.column_combo = ttk.Combobox(options, textvariable=self.column_var, state="readonly")
        self.column_combo.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._select_sheet)
        self.column_combo.bind("<<ComboboxSelected>>", self._select_column)
        ttk.Label(body, text="파일명 패턴 · %는 분류값으로 바뀝니다", style="Note.TLabel").grid(sticky="w", pady=(8, 4))
        self.pattern_entry = ttk.Entry(body, textvariable=self.pattern_var)
        self.pattern_entry.grid(sticky="ew", ipady=2)
        self._input_widgets.extend((self.sheet_combo, self.column_combo, self.pattern_entry))
        body = self._section(page, 2, "3   저장 위치")
        self.output_entry, _ = self._path_row(body, self.output_var, self._browse_output, "출력 폴더")
        self.split_path_note = tk.StringVar()
        self._wrapped_label(body, textvariable=self.split_path_note, style="Note.TLabel").grid(sticky="ew")
        preview_frame = ttk.Frame(body)
        preview_frame.grid(sticky="ew", pady=(10, 0))
        preview_frame.columnconfigure(0, weight=1)
        self.preview_tree = ttk.Treeview(preview_frame, columns=("label", "count", "filename"), show="headings", height=4)
        for col, title, width in (("label", "분류", 180), ("count", "행 수", 80), ("filename", "출력 파일", 560)):
            self.preview_tree.heading(col, text=title)
            self.preview_tree.column(col, width=width)
        self.preview_tree.grid(row=0, column=0, sticky="ew")
        bar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        bar.grid(row=0, column=1, sticky="ns")
        self.preview_tree.configure(yscrollcommand=bar.set)
        preview_frame.grid_remove()
        self._work_footer(page, "split", "분할 시작", self._split)
        self.progress = self.split_progress
        self.status_var = self.split_status_var

    def _build_merge(self):
        page = self._page("파일 병합")
        body = self._section(page, 0, "1   병합할 파일")
        actions = ttk.Frame(body)
        actions.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.merge_list_buttons = []
        for i, (label, action) in enumerate((("파일 추가", self._add_merge_files), ("선택 제거", self._remove_merge_file),
                                            ("목록 비우기", self._clear_merge_files),
                                            ("위로", lambda: self._move_merge_file(-1)), ("아래로", lambda: self._move_merge_file(1)))):
            button = ttk.Button(actions, text=label, command=action)
            button.grid(row=0, column=i, padx=(0, 6))
            self._input_widgets.append(button)
            self.merge_list_buttons.append(button)
        actions.columnconfigure(5, weight=1)
        self.merge_count_var = tk.StringVar(value="파일 0개")
        ttk.Label(actions, textvariable=self.merge_count_var, style="Note.TLabel").grid(row=0, column=5, sticky="e")
        self.merge_tree = ttk.Treeview(body, columns=("path", "sheet", "rows", "status"), show="headings", height=6, selectmode="extended")
        for col, title, width in (("path", "파일 이름 · 위에서 아래 순서로 병합", 390), ("sheet", "시트", 130), ("rows", "데이터 행", 90), ("status", "상태", 110)):
            self.merge_tree.heading(col, text=title, anchor="w")
            self.merge_tree.column(col, width=width, minwidth=70, stretch=col == "path", anchor="e" if col == "rows" else "w")
        self.merge_tree.grid(row=1, column=0, sticky="ew")
        bar = ttk.Scrollbar(body, orient="vertical", command=self.merge_tree.yview)
        bar.grid(row=1, column=1, sticky="ns")
        hbar = ttk.Scrollbar(body, orient="horizontal", command=self.merge_tree.xview)
        hbar.grid(row=2, column=0, sticky="ew")
        self.merge_tree.configure(yscrollcommand=bar.set, xscrollcommand=hbar.set)
        self.merge_selected_path_var = tk.StringVar()
        self._wrapped_label(body, textvariable=self.merge_selected_path_var, style="Note.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(3, 0))
        self.merge_tree.bind("<<TreeviewSelect>>", self._merge_selection_changed)
        self.merge_tree.bind("<Delete>", lambda _: self._remove_merge_file())
        body = self._section(page, 1, "2   병합 방식")
        self._wrapped_label(body, text="같은 열 이름·순서의 Excel 표(Table)를 합칩니다. 첫 파일의 서식을 사용하고 표 데이터 본문 수식은 현재 계산값으로 저장합니다.",
                            style="Note.TLabel").grid(sticky="ew")
        self.merge_output_var = tk.StringVar()
        self._output_section(page, "merge", self.merge_output_var, self._browse_merge_output)
        self._work_footer(page, "merge", "병합 시작", self._merge)

    def _build_etc(self) -> None:
        page = self._page("시트 정리")
        self.etc_source_var, self.etc_sheet_var, self.etc_output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.etc_remove_artifacts_var = tk.BooleanVar(value=False)
        self.etc_reset_fill_var = tk.BooleanVar(value=False)
        self.etc_exclude_table_headers_var = tk.BooleanVar(value=True)
        self.etc_remove_conditional_formats_var = tk.BooleanVar(value=False)
        body = self._section(page, 0, "1   파일")
        self._path_row(body, self.etc_source_var, self._browse_etc_source, "원본 파일", readonly=True)
        body = self._section(page, 1, "2   작업 설정")
        ttk.Label(body, text="워크시트").grid(sticky="w")
        self.etc_sheet_combo = ttk.Combobox(body, textvariable=self.etc_sheet_var, state="disabled")
        self.etc_sheet_combo.grid(sticky="ew", pady=6)
        self.etc_sheet_combo.bind("<<ComboboxSelected>>", lambda _: self._render_etc_state())
        self.etc_remove_artifacts_checkbox = ttk.Checkbutton(body, text="도형·메모·댓글 삭제 (그림, 차트, 버튼 포함)",
            variable=self.etc_remove_artifacts_var, command=self._render_etc_state)
        self.etc_remove_artifacts_checkbox.grid(sticky="w", pady=6)
        fill = ttk.Frame(body)
        fill.grid(sticky="w", pady=6)
        self.etc_reset_fill_checkbox = ttk.Checkbutton(fill, text="채우기 색 초기화", variable=self.etc_reset_fill_var, command=self._render_etc_state)
        self.etc_reset_fill_checkbox.grid(row=0, column=0, sticky="w")
        self.etc_exclude_table_headers_checkbox = ttk.Checkbutton(fill, text="테이블 헤더 제외", variable=self.etc_exclude_table_headers_var, command=self._render_etc_state)
        self.etc_exclude_table_headers_checkbox.grid(row=0, column=1, sticky="w", padx=16)
        self.etc_remove_conditional_formats_checkbox = ttk.Checkbutton(fill, text="조건부 서식 삭제", variable=self.etc_remove_conditional_formats_var, command=self._render_etc_state)
        self.etc_remove_conditional_formats_checkbox.grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Label(body, text="채우기 색은 직접 지정한 색만 초기화합니다. 조건부 서식 삭제는 헤더를 포함한 시트 전체에 적용합니다.",
                  style="Note.TLabel", wraplength=850).grid(sticky="w", pady=(6, 0))
        self._input_widgets.extend((self.etc_sheet_combo, self.etc_remove_artifacts_checkbox, self.etc_reset_fill_checkbox,
                                    self.etc_exclude_table_headers_checkbox, self.etc_remove_conditional_formats_checkbox))
        self._output_section(page, "etc", self.etc_output_var, self._browse_etc_output)
        self._work_footer(page, "etc", "정리 시작", self._run_etc)

    def _build_error_panel(self):
        self.error_panel = ttk.Frame(self.root, padding=(20, 10))
        self.error_panel.grid(row=2, column=0, sticky="ew")
        self.error_panel.columnconfigure(0, weight=1)
        self.error_message_var = tk.StringVar()
        self._wrapped_label(self.error_panel, textvariable=self.error_message_var, foreground="#a3372b").grid(row=0, column=0, sticky="ew")
        actions = ttk.Frame(self.error_panel)
        actions.grid(row=1, column=0, sticky="w", pady=6)
        for i, (label, action) in enumerate((("상세 내용", self._toggle_error_detail), ("오류 복사", self._copy_error),
                                            ("로그 폴더 열기", lambda: self._open_path(_log_path(self.logger).parent)))):
            ttk.Button(actions, text=label, command=action).grid(row=0, column=i, padx=(0, 6))
        self.error_text = tk.Text(self.error_panel, height=4, wrap="word", state="disabled")
        self.error_text.grid(row=2, column=0, sticky="ew")
        self.error_text.grid_remove()
        self.error_panel.grid_remove()

    def _render_etc_state(self) -> None:
        error = self._output_error("etc", (self.etc_source_var.get(),))
        ready = all((self.etc_source_var.get(), self.etc_sheet_var.get(), self.etc_output_var.get()))
        ready = ready and (self.etc_remove_artifacts_var.get() or self.etc_reset_fill_var.get()
                           or self.etc_remove_conditional_formats_var.get())
        self.etc_reason_var.set("작업 중에는 설정을 바꿀 수 없습니다." if self._busy else
                                error or ("" if ready else "파일, 시트와 정리할 작업을 선택하세요."))
        ready = ready and not error
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
        if self._busy or self._output_error("etc", (self.etc_source_var.get(),)):
            return
        source, target = Path(self.etc_source_var.get()), Path(self.etc_output_var.get())
        sheet_name = self.etc_sheet_var.get()
        remove_artifacts, reset_fill = self.etc_remove_artifacts_var.get(), self.etc_reset_fill_var.get()
        remove_conditional_formats = self.etc_remove_conditional_formats_var.get()
        exclude_table_headers = self.etc_exclude_table_headers_var.get()
        self._start_worker(lambda: ("etc_execute", self.etc_service.execute(
            source, sheet_name, target, remove_artifacts=remove_artifacts, reset_fill=reset_fill,
            exclude_table_headers=exclude_table_headers,
            remove_conditional_formats=remove_conditional_formats,
            progress=lambda completed, total, label: self.events.put(("progress", completed, total, label)),
        )), execution=True)

    def _build_compare(self) -> None:
        page = self._page("변경 비교")
        self.compare_reference_var, self.compare_comparison_var, self.compare_output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        body = self._section(page, 0, "1   파일")
        body.columnconfigure((0, 1), weight=1)
        for col, (label, variable, kind) in enumerate((("기준 파일", self.compare_reference_var, "reference"), ("비교대상 파일", self.compare_comparison_var, "comparison"))):
            card = ttk.Frame(body, padding=10, relief="solid", borderwidth=1)
            card.grid(row=0, column=col, sticky="ew", padx=(0, 8) if col == 0 else (0, 0))
            card.columnconfigure(0, weight=1)
            self._path_row(card, variable, lambda k=kind: self._browse_compare_input(k), label, readonly=True)
        body = self._section(page, 1, "2   작업 설정")
        modes = ttk.Frame(body)
        modes.grid(row=0, column=0, sticky="ew")
        modes.columnconfigure((0, 1), weight=1)
        self.compare_by_key_var = tk.BooleanVar(value=False)
        for col, (name, description, value, attr) in enumerate((
            ("셀 위치로 비교", "같은 시트·행·열의 값을 비교합니다.", False, "compare_position_radio"),
            ("키가 같은 행 비교", "행 순서가 달라도 선택한 키로 찾습니다.", True, "compare_key_radio"),
        )):
            card = ttk.Frame(modes, padding=10, relief="solid", borderwidth=1)
            card.grid(row=0, column=col, sticky="ew", padx=(0, 8) if col == 0 else (0, 0))
            radio = ttk.Radiobutton(card, text=name, variable=self.compare_by_key_var, value=value, command=self._compare_mode_changed)
            radio.grid(sticky="w")
            ttk.Label(card, text=description, style="Note.TLabel").grid(sticky="w", padx=(22, 0), pady=(5, 0))
            setattr(self, attr, radio)
            self._input_widgets.append(radio)
        options = self.compare_key_options = ttk.Frame(body)
        options.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        options.columnconfigure((0, 1), weight=1)
        self.compare_tables_button = ttk.Button(options, text="표 불러오기", command=self._load_compare_tables)
        self.compare_tables_button.grid(row=0, column=1, sticky="e", pady=(0, 8))
        self.compare_reference_table_combo = ttk.Combobox(options, state="disabled")
        self.compare_comparison_table_combo = ttk.Combobox(options, state="disabled")
        for col, (label, combo) in enumerate((("기준 표", self.compare_reference_table_combo), ("비교대상 표", self.compare_comparison_table_combo))):
            ttk.Label(options, text=label).grid(row=1, column=col, sticky="w")
            combo.grid(row=2, column=col, sticky="ew", padx=(0, 8) if col == 0 else (0, 0), pady=5)
            combo.bind("<<ComboboxSelected>>", self._refresh_compare_keys)
        ttk.Label(options, text="키 컬럼 · 여러 개 선택할 수 있습니다", style="Note.TLabel").grid(row=3, column=0, columnspan=2, sticky="w")
        self.compare_key_list = tk.Listbox(options, selectmode="multiple", exportselection=False, height=3, relief="solid", borderwidth=1)
        self.compare_key_list.grid(row=4, column=0, columnspan=2, sticky="ew", pady=5)
        bar = ttk.Scrollbar(options, orient="vertical", command=self.compare_key_list.yview)
        bar.grid(row=4, column=2, sticky="ns")
        self.compare_key_list.configure(yscrollcommand=bar.set)
        self.compare_key_list.bind("<<ListboxSelect>>", lambda _: self._render_compare_state())
        self._input_widgets.extend((self.compare_tables_button, self.compare_reference_table_combo, self.compare_comparison_table_combo, self.compare_key_list))
        self._output_section(page, "compare", self.compare_output_var, self._browse_compare_output)
        self._work_footer(page, "compare", "비교 시작", self._compare)

    def _render_compare_state(self) -> None:
        error = self._output_error("compare", (self.compare_reference_var.get(), self.compare_comparison_var.get()))
        ready = all(variable.get() for variable in (
            self.compare_reference_var, self.compare_comparison_var, self.compare_output_var,
        ))
        key_mode = self.compare_by_key_var.get()
        if key_mode:
            self.compare_key_options.grid()
        else:
            self.compare_key_options.grid_remove()
        enabled = key_mode and not self._busy
        has_sources = bool(self.compare_reference_var.get() and self.compare_comparison_var.get())
        self.compare_tables_button.configure(state="normal" if enabled and has_sources else "disabled")
        for combo, tables in zip((self.compare_reference_table_combo, self.compare_comparison_table_combo), self.compare_tables):
            combo.configure(state="readonly" if enabled and tables else "disabled")
        self.compare_key_list.configure(state="normal" if enabled and self.compare_key_list.size() else "disabled")
        if key_mode:
            ready = ready and bool(self.compare_key_list.curselection())
        self.compare_reason_var.set("작업 중에는 설정을 바꿀 수 없습니다." if self._busy else
                                    error or ("" if ready else "기준·비교대상 파일을 선택하고, 키 비교라면 표와 키를 선택하세요."))
        ready = ready and not error
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
        self.compare_status_var.set("파일과 결과 저장 위치를 확인하고 비교 시작을 누르세요.")
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
        if self._busy or self._output_error("compare", (self.compare_reference_var.get(), self.compare_comparison_var.get())):
            return
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
        self._compare_table_names = (
            options["reference_table"][0], options["comparison_table"][0]
        ) if options else None
        self._start_worker(lambda: ("compare_execute", self.compare_service.execute(
            reference, comparison, target, **options,
            progress=lambda completed, total, label: self.events.put(("progress", completed, total, label)),
        )), execution=True)

    def _render_merge_state(self) -> None:
        error = self._output_error("merge", self.merge_sources)
        ready = len(self.merge_sources) >= 2 and bool(self.merge_output_var.get())
        self.merge_reason_var.set("작업 중에는 설정을 바꿀 수 없습니다." if self._busy else
                                  error or ("시작하면 파일을 검사하고 병합합니다." if ready else "파일 두 개 이상을 추가하세요."))
        ready = ready and not error
        self.merge_button.configure(state="normal" if ready and not self._busy else "disabled")
        self._merge_selection_changed()

    def _merge_selection_changed(self, _event=None) -> None:
        selection = self.merge_tree.selection()
        self.merge_selected_path_var.set(str(self.merge_sources[int(selection[0])]) if selection else "파일 추가로 병합할 파일을 선택하세요. Ctrl / Shift로 여러 파일을 선택할 수 있습니다.")
        count = len(self.merge_sources)
        self.merge_count_var.set(f"파일 {count:,}개")
        enabled = not self._busy
        one = len(selection) == 1
        for button, available in zip(self.merge_list_buttons[1:], (
            bool(selection), bool(count), one and int(selection[0]) > 0, one and int(selection[0]) < count - 1,
        )):
            button.configure(state="normal" if enabled and available else "disabled")

    def _output_error(self, kind: str, sources) -> str:
        value = getattr(self, kind + "_output_var").get()
        error = validate_output_path(value, sources=sources, allow_existing=kind == "merge")
        note = error or "원본을 유지하고 결과를 새 파일로 저장합니다."
        if not error and self.result_paths.get(kind):
            note = "다음 실행에 사용할 파일명입니다. 저장된 파일은 아래 결과 패널에서 여세요."
        if not error and kind == "merge" and Path(value).exists():
            note = "기존 결과 파일입니다. 병합 전에 덮어쓰기를 확인합니다. 새 이름을 권장합니다."
        getattr(self, kind + "_path_note").set(note)
        return error

    def _recommend_output(self, kind: str) -> None:
        variable = getattr(self, kind + "_output_var")
        if kind == "merge":
            base = self.merge_sources[0].parent / "merged.xlsx" if self.merge_sources else Path.cwd() / "merged.xlsx"
        else:
            source = self.compare_comparison_var.get() if kind == "compare" else self.etc_source_var.get()
            path = Path(source) if source else Path.cwd() / "결과.xlsx"
            base = path.with_name(path.stem + ("_비교결과.xlsx" if kind == "compare" else "_정리결과.xlsx"))
        current = Path(variable.get()) if variable.get() else base
        # Invalid extensions/parents should not make the recommendation unusable.
        if current.suffix.lower() != ".xlsx" or not current.parent.is_dir():
            current = base
        try:
            variable.set(str(suggest_output_path(current)))
        except OSError as exc:
            self._handle_error(exc)

    def _use_merge_for_compare(self) -> None:
        if self._busy or not self.result_paths.get("merge"):
            return
        path = self.result_paths["merge"][0]
        self.compare_comparison_var.set(str(path))
        self._invalidate_compare_tables()
        self.compare_output_var.set(str(suggest_output_path(path.with_name(path.stem + "_비교결과.xlsx"))))
        self.notebook.select(2)
        self.compare_status_var.set("병합 결과를 비교대상으로 넣었습니다. 기준 파일과 비교 방식을 확인하세요.")
        self._render_compare_state()

    def _invalidate_merge_preview(self) -> None:
        self.merge_preview = None
        self.merge_tree.delete(*self.merge_tree.get_children())
        for index, path in enumerate(self.merge_sources):
            self.merge_tree.insert("", "end", iid=str(index), values=(path.name, "", "", "검사 전"))
        if self.merge_sources:
            self.merge_tree.selection_set("0")
        self.merge_selected_path_var.set(str(self.merge_sources[0]) if self.merge_sources else "")
        self.merge_status_var.set(f"파일 {len(self.merge_sources)}개 · 순서를 확인하고 병합 시작을 누르세요.")
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
            self.merge_output_var.set(str(suggest_output_path(self.merge_sources[0].parent / "merged.xlsx")))
        self._invalidate_merge_preview()

    def _remove_merge_file(self) -> None:
        if self._busy:
            return
        selection = self.merge_tree.selection()
        if selection:
            for index in sorted(map(int, selection), reverse=True):
                del self.merge_sources[index]
            self._invalidate_merge_preview()

    def _clear_merge_files(self) -> None:
        if not self._busy:
            self.merge_sources.clear()
            self._invalidate_merge_preview()

    def _move_merge_file(self, step: int) -> None:
        selection = self.merge_tree.selection()
        if self._busy or len(selection) != 1:
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
        if self._busy or self._output_error("merge", self.merge_sources):
            return
        sources = tuple(self.merge_sources)
        target = Path(self.merge_output_var.get())
        self._invalidate_merge_preview()
        self._start_worker(lambda: ("merge_preview", self.merge_service.preview(sources, target)))
        self._phase = "병합 파일 검사 중"

    def _merge(self) -> None:
        if self._busy or len(self.merge_sources) < 2 or self._output_error("merge", self.merge_sources):
            return
        self._merge_after_preview = True
        self._preview_merge()

    def _execute_merge(self) -> None:
        preview = self.merge_preview
        if self._busy or preview is None or self._output_error("merge", self.merge_sources):
            return
        if preview.prior_signature is not None:
            prompt = f"기존 결과 파일을 덮어씁니다. 계속할까요?\n\n{preview.target}"
            if not messagebox.askyesno("결과 덮어쓰기", prompt, parent=self.root):
                self.merge_status_var.set("덮어쓰기를 취소했습니다. 새 파일명을 지정하고 시작하세요.")
                return
        self._start_worker(lambda: ("merge_execute", self.merge_service.execute(
            preview, preview.prior_signature is not None,
            lambda completed, total, label: self.events.put(("progress", completed, total, label)),
        )), execution=True)

    def _set_busy(self, busy: bool) -> None:
        index = self.notebook.index(self.notebook.select())
        self.progress_var = self._progress_variables[index]
        self.status_var = (self.split_status_var, self.merge_status_var, self.compare_status_var, self.etc_status_var)[index]
        super()._set_busy(busy)
        selected = self.notebook.select()
        for tab in self.notebook.tabs():
            self.notebook.tab(tab, state="disabled" if busy and tab != selected else "normal")
        if not busy:
            for kind, (_, _, _, buttons) in self.result_panels.items():
                for button in buttons:
                    button.configure(state="normal" if self.result_paths.get(kind) else "disabled")
        self._render_merge_state()
        self._render_compare_state()
        self._render_etc_state()
        if busy and index == 1:
            for item in self.merge_tree.get_children():
                self.merge_tree.set(item, "status", "병합 대기" if self._executing else "검사 중")

    def _handle_ok(self, payload: object) -> None:
        tag, value = payload
        if tag == "etc_source":
            self.etc_sheet_combo.configure(values=value)
            self.etc_sheet_var.set(value[0] if value else "")
            self._set_busy(False)
            self.etc_status_var.set("정리할 시트와 작업을 선택하세요.")
        elif tag == "etc_execute":
            self._set_busy(False)
            self.etc_status_var.set("정리 완료 · 결과를 저장했습니다.")
            self._show_result("etc", (Path(value),), "시트 정리 완료")
            self.etc_output_var.set(str(suggest_output_path(value)))
        elif tag == "compare_tables":
            self.compare_tables = value
            for combo, tables in zip((self.compare_reference_table_combo, self.compare_comparison_table_combo), value):
                combo.configure(values=tuple(f"{table.sheet_name} / {table.table_name}" for table in tables))
                if tables:
                    combo.current(0)
                else:
                    combo.set("")
            self._refresh_compare_keys()
            self._set_busy(False)
            self.compare_status_var.set("양쪽 표를 확인하고 키 컬럼을 선택하세요. 키 중복·누락은 실행 시 검사합니다.")
        elif tag == "merge_preview":
            self.merge_preview = value
            for index, item in enumerate(value.inputs):
                self.merge_tree.item(str(index), values=(item.source.name, item.sheet_name, f"{item.row_count:,}", "준비 완료"))
            self._set_busy(False)
            self.merge_status_var.set(f"파일 {len(value.inputs)}개 · 데이터 {value.row_count}행 · 병합 준비 완료")
            if self._merge_after_preview:
                self._merge_after_preview = False
                self._execute_merge()
        elif tag == "merge_execute":
            self._set_busy(False)
            self._show_result("merge", (Path(value),), "병합 완료")
            self.merge_output_var.set(str(suggest_output_path(value)))
            for item in self.merge_tree.get_children():
                self.merge_tree.set(item, "status", "병합 완료")
            self.merge_status_var.set("병합 완료 · 결과 파일을 열거나 다음 작업을 준비하세요.")
        elif tag == "compare_execute":
            self._set_busy(False)
            modified = getattr(value, "modified_cells", value.changed_cells)
            added = getattr(value, "added_rows", 0)
            if self.compare_by_key_var.get():
                summary = f"비교 완료 · 값 변경 {modified}셀 · 추가 {added}행 · 누락 {value.missing_rows}행"
                summary += f" · 노란색 표시 {value.changed_cells}셀"
            else:
                summary = f"위치 비교 완료 · 다른 값 {value.changed_cells}셀을 노란색으로 표시했습니다."
            if value.missing_sheets:
                summary += "\n누락 시트: " + ", ".join(value.missing_sheets)
            if value.missing_columns:
                summary += "\n누락 열: " + ", ".join(value.missing_columns)
            if getattr(value, "details_truncated", False):
                summary += (f"\n상세 목록은 처음 1,000건만 표시합니다 ({value.omitted_details}건 생략)."
                            " 요약·색칠은 전체 비교 기준입니다. 누락 행·시트는 결과 파일에 추가되지 않습니다.")
            labels = {"changed": "값 변경", "added": "추가", "missing": "누락"}
            rows = []
            for detail in getattr(value, "details", ()):
                coordinates = []
                table_names = getattr(self, "_compare_table_names", None)
                if detail.reference_cell:
                    coordinates.append("기준 " + ((table_names[0] + "!") if table_names else "") + detail.reference_cell)
                if detail.comparison_cell:
                    coordinates.append("대상 " + ((table_names[1] + "!") if table_names else "") + detail.comparison_cell)
                location = " · ".join(part for part in (detail.sheet_name, detail.key, ", ".join(coordinates) or detail.cell) if part)
                rows.append((labels.get(detail.kind, detail.kind), location, detail.column_name,
                             "" if detail.reference_value is None else str(detail.reference_value),
                             "" if detail.comparison_value is None else str(detail.comparison_value)))
            self.compare_status_var.set(summary)
            self._show_result("compare", (value.target,), summary, rows)
            self.compare_output_var.set(str(suggest_output_path(value.target)))
        else:
            super()._handle_ok(payload)

    def _handle_error(self, error: object) -> None:
        self._merge_after_preview = False
        selected = self.notebook.index(self.notebook.select())
        if selected == 1:
            self._invalidate_merge_preview()
            for item in self.merge_tree.get_children():
                self.merge_tree.set(item, "status", "다시 검사 필요")
        elif selected == 0:
            self.controller.set_pattern(self.pattern_var.get())
            self._clear_preview()
        super()._handle_error(error)
        status = (self.status_var, self.merge_status_var, self.compare_status_var, self.etc_status_var)[selected]
        status.set("오류가 발생했습니다. 아래 안내를 확인하고 다시 실행하세요.")

    def _show_progress(self, completed: int, total: int, label: str) -> None:
        super()._show_progress(completed, total, label)
        if (self._executing and self.notebook.index(self.notebook.select()) == 1
                and total == len(self.merge_sources) and 1 <= completed <= total):
            self.merge_tree.set(str(completed - 1), "status", "데이터 반영")

    def _progress_widget(self) -> ttk.Progressbar:
        return (self.progress, self.merge_progress, self.compare_progress, self.etc_progress)[
            self.notebook.index(self.notebook.select())
        ]
