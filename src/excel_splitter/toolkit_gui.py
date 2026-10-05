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
        self._pages = {}
        self._edit_areas = {}
        self._save_rows = {}
        self._result_visible = set()
        self.result_summaries = {}
        self.result_back_buttons = {}
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
        self.root.geometry("900x600")
        self.root.minsize(900, 600)
        self.root.option_add("*Font", ("맑은 고딕", 10))
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self.root.configure(background="#f4f6f9")
        style.configure(".", font=("맑은 고딕", 10), background="#ffffff", foreground="#1d2735")
        style.configure("TFrame", background="#ffffff")
        style.configure("TLabel", background="#ffffff")
        style.configure("Shell.TFrame", background="#f4f6f9")
        style.configure("Shell.TLabel", background="#f4f6f9", foreground="#627086")
        style.configure("Note.TLabel", foreground="#627086")
        style.configure("Title.TLabel", font=("맑은 고딕", 12, "bold"))
        style.configure("TButton", padding=(10, 5), background="#ffffff", bordercolor="#dce2eb")
        style.configure("Nav.TButton", padding=(18, 8), borderwidth=0)
        style.configure("Selected.Nav.TButton", background="#edf3fc", foreground="#1b60bc")
        style.configure("Primary.TButton", background="#1b60bc", foreground="white", padding=(18, 10))
        style.map("Primary.TButton", background=[("disabled", "#e4eaf0"), ("active", "#164f9a")],
                  foreground=[("disabled", "#718096")])
        style.configure("TEntry", padding=4, fieldbackground="#ffffff", bordercolor="#dce2eb")
        style.map("TEntry", fieldbackground=[("readonly", "#f4f6f9")])
        style.configure("TCombobox", padding=4, fieldbackground="#ffffff", bordercolor="#dce2eb")
        style.map("TCombobox", fieldbackground=[("readonly", "#ffffff")])
        style.configure("Treeview", rowheight=32, fieldbackground="#ffffff", bordercolor="#dce2eb")
        style.configure("Treeview.Heading", padding=(8, 6), background="#f4f6f9", foreground="#627086")
        style.map("Treeview", background=[("selected", "#edf3fc")], foreground=[("selected", "#1b60bc")])
        style.configure("Horizontal.TProgressbar", troughcolor="#e7ecf3", background="#1b60bc", thickness=5)
        style.configure("Workspace.TNotebook", borderwidth=0, tabmargins=0)
        style.layout("Workspace.TNotebook.Tab", [])
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)
        header = ttk.Frame(self.root, padding=(16, 7), style="Shell.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Excel 도구", style="Shell.TLabel", foreground="#1d2735",
                  font=("맑은 고딕", 11, "bold")).pack(side="left")
        ttk.Label(header, text="원본 보존 · 검증 후 결과 저장", style="Shell.TLabel").pack(side="right")
        nav = ttk.Frame(self.root, padding=(10, 4))
        nav.grid(row=1, column=0, sticky="ew")
        self.notebook = ttk.Notebook(self.root, style="Workspace.TNotebook")
        self.notebook.grid(row=2, column=0, sticky="nsew")
        self.nav_buttons = {}
        for kind, label, index in (("merge", "병합", 1), ("split", "분할", 0), ("compare", "비교", 2), ("etc", "시트 정리", 3)):
            button = ttk.Button(nav, text=label, style="Nav.TButton", command=lambda i=index: self.notebook.select(i))
            button.pack(side="left", padx=(0, 4))
            self.nav_buttons[kind] = button
            self._input_widgets.append(button)
        ttk.Button(nav, text="로그", command=self._show_log).pack(side="right")
        self._build_split_page(self._page("분할", "split"))
        self._build_merge()
        self._build_compare()
        self._build_etc()
        self._build_error_panel()
        self.notebook.bind("<<NotebookTabChanged>>", self._tab_changed)
        self.notebook.select(1)
        self._tab_changed()

    def _tab_changed(self, _event=None):
        current = self.notebook.index(self.notebook.select())
        for index, kind in enumerate(("split", "merge", "compare", "etc")):
            self.nav_buttons[kind].configure(style="Selected.Nav.TButton" if index == current else "Nav.TButton")

    def _page(self, title: str, kind: str):
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text=title)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        workspace = ttk.Frame(outer)
        workspace.grid(row=0, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(0, weight=1)
        editor = ttk.Frame(workspace)
        editor.grid(row=0, column=0, sticky="nsew")
        editor.columnconfigure(0, weight=1)
        editor.rowconfigure(2, weight=1)
        self._pages[kind] = outer
        self._edit_areas[kind] = editor
        return editor

    def _ribbon(self, page):
        ribbon = ttk.Frame(page, padding=(16, 10), style="Shell.TFrame")
        ribbon.grid(row=0, column=0, sticky="ew")
        return ribbon

    def _field(self, parent, column: int, title: str):
        parent.columnconfigure(column, weight=1, uniform="settings")
        field = ttk.Frame(parent, style="Shell.TFrame")
        field.grid(row=0, column=column, sticky="ew", padx=(0, 12))
        field.columnconfigure(0, weight=1)
        ttk.Label(field, text=title, style="Shell.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        return field

    def _path_row(self, parent, variable, action, label="", readonly=False):
        row = ttk.Frame(parent)
        row.grid(sticky="ew")
        row.columnconfigure(0, weight=1)
        caption = ttk.Frame(row)
        caption.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))
        caption.columnconfigure(1, weight=1)
        ttk.Label(caption, text=label, style="Note.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
        if readonly:
            name = ttk.Label(caption, width=1, anchor="e", font=("맑은 고딕", 10, "bold"))
            name.grid(row=0, column=1, sticky="ew")
            self.source_name_labels[str(variable)] = name
            variable.trace_add("write", lambda *_: name.configure(text=Path(variable.get()).name if variable.get() else "선택한 파일 없음"))
        entry = ttk.Entry(row, textvariable=variable, state="readonly" if readonly else "normal", width=1)
        entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        browse = ttk.Button(row, text="찾아보기", command=action)
        browse.grid(row=1, column=1)
        ttk.Button(row, text="경로", command=lambda: self._show_text("전체 경로", variable.get() or "선택한 경로가 없습니다.")).grid(row=1, column=2, padx=(6, 0))
        entry.bind("<Double-1>", lambda _: self._show_text("전체 경로", variable.get()))
        self._input_widgets.append(browse)
        if not readonly:
            self._input_widgets.append(entry)
        return entry, row

    def _tree(self, parent, columns, *, selectmode="browse"):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        tree = ttk.Treeview(parent, columns=tuple(item[0] for item in columns), show="headings", height=1, selectmode=selectmode)
        for name, title, width in columns:
            tree.heading(name, text=title, anchor="w")
            tree.column(name, width=width, minwidth=65, stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        return tree

    def _work_footer(self, page, kind, text, action, variable, browse):
        footer = ttk.Frame(self._pages[kind], padding=(16, 9), style="Shell.TFrame")
        footer.grid(row=1, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        save = ttk.Frame(footer, style="Shell.TFrame")
        save.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        save.columnconfigure(1, weight=1)
        ttk.Label(save, text="출력 폴더" if kind == "split" else "저장할 파일", style="Shell.TLabel").grid(row=0, column=0, padx=(0, 9))
        entry = ttk.Entry(save, textvariable=variable, width=1)
        entry.grid(row=0, column=1, sticky="ew")
        entry.bind("<Double-1>", lambda _: self._show_text("저장 경로", variable.get()))
        choose = ttk.Button(save, text="찾아보기", command=browse)
        choose.grid(row=0, column=2, padx=(6, 0))
        self._input_widgets.extend((entry, choose))
        if kind != "split":
            recommend = ttk.Button(save, text="새 이름", command=lambda: self._recommend_output(kind))
            recommend.grid(row=0, column=3, padx=(6, 0))
            self._input_widgets.append(recommend)
        self._save_rows[kind] = save
        setattr(self, "output_entry" if kind == "split" else kind + "_output_entry", entry)
        status, reason, note = tk.StringVar(value="시작하면 자동 검사 후 실행합니다."), tk.StringVar(), tk.StringVar()
        setattr(self, kind + "_status_var", status)
        setattr(self, kind + "_reason_var", reason)
        setattr(self, kind + "_path_note", note)
        status_label = ttk.Label(footer, width=1, style="Shell.TLabel")
        status_label.grid(row=2, column=0, sticky="ew", padx=(0, 12), pady=(4, 0))
        def update_status(*_):
            value = status.get() if self._busy or kind in self._result_visible else reason.get() or status.get()
            status_label.configure(text=value.splitlines()[0] if value else "")
        status.trace_add("write", update_status)
        reason.trace_add("write", update_status)
        update_status()
        status_label.bind("<Double-1>", lambda _: self._show_text("작업 상태", "\n\n".join(value.get() for value in (status, reason, note) if value.get())))
        progress_var = tk.DoubleVar(value=0)
        self._progress_variables.append(progress_var)
        progress = ttk.Progressbar(footer, variable=progress_var, maximum=1)
        progress.grid(row=1, column=0, sticky="ew", padx=(0, 12), pady=(7, 0))
        setattr(self, kind + "_progress", progress)
        button = ttk.Button(footer, text=text, style="Primary.TButton", command=lambda: self._show_editor(kind) if kind in self._result_visible else action())
        button.grid(row=0, column=1, rowspan=3, sticky="ns")
        self._input_widgets.append(button)
        setattr(self, kind + "_button", button)
        self._build_result_panel(page, kind)

    def _build_result_panel(self, page, kind):
        frame = ttk.Frame(page.master, padding=(16, 12))
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        label = ttk.Label(frame, width=1, style="Title.TLabel")
        label.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        actions = ttk.Frame(frame)
        actions.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        buttons = []
        for title, action in (("선택 결과 열기" if kind == "split" else "결과 열기", lambda: self._open_result(kind)),
                              ("폴더 열기", lambda: self._open_result(kind, folder=True))):
            button = ttk.Button(actions, text=title, command=action)
            button.pack(side="left", padx=(0, 6))
            buttons.append(button)
        if kind == "merge":
            button = ttk.Button(actions, text="이 결과 비교", command=self._use_merge_for_compare)
            button.pack(side="left", padx=(0, 6))
            buttons.append(button)
        back = ttk.Button(actions, text="설정으로 돌아가기", command=lambda: self._show_editor(kind))
        back.pack(side="right")
        self.result_back_buttons[kind] = back
        ttk.Button(actions, text="전체 요약", command=lambda: self._show_text("작업 결과", self.result_summaries.get(kind, ""))).pack(side="right", padx=6)
        ttk.Button(actions, text="선택 상세", command=lambda: self._show_result_detail(kind)).pack(side="right")
        table = ttk.Frame(frame)
        table.grid(row=2, column=0, sticky="nsew")
        tree = self._tree(table, (("kind", "구분", 80), ("location", "시트 · 키 / 위치", 190),
                                  ("column", "열", 110), ("reference", "기준 파일", 170), ("comparison", "대상 파일", 170)))
        if kind != "compare":
            tree.configure(displaycolumns=("kind", "location", "comparison"))
            tree.heading("location", text="결과 파일")
            tree.heading("comparison", text="저장 경로 / 오류")
        tree.bind("<Double-1>", lambda _: self._show_result_detail(kind))
        self._input_widgets.extend((*buttons, back))
        self.result_panels[kind] = (frame, label, tree, buttons)
        frame.grid_remove()

    def _show_editor(self, kind):
        if self._busy:
            return
        self._result_visible.discard(kind)
        self.result_panels[kind][0].grid_remove()
        self._edit_areas[kind].grid()
        self._save_rows[kind].grid()
        getattr(self, kind + "_button").configure(text={"split": "분할 시작", "merge": "병합 시작", "compare": "비교 시작", "etc": "정리 시작"}[kind])
        self._render_state(self.controller.state)
        self._render_merge_state()
        self._render_compare_state()
        self._render_etc_state()

    def _render_state(self, state):
        super()._render_state(state)
        if "split" in self._result_visible and not self._busy:
            self.split_button.configure(state="normal")

    def _build_split_page(self, page):
        ribbon = self._ribbon(page)
        self.sheet_combo = ttk.Combobox(self._field(ribbon, 0, "워크시트"), textvariable=self.sheet_var, state="disabled", width=1)
        self.column_combo = ttk.Combobox(self._field(ribbon, 1, "분류 열"), textvariable=self.column_var, state="disabled", width=1)
        self.pattern_entry = ttk.Entry(self._field(ribbon, 2, "파일명 패턴 · % = 분류값"), textvariable=self.pattern_var, width=1)
        for widget in (self.sheet_combo, self.column_combo, self.pattern_entry):
            widget.grid(row=1, column=0, sticky="ew")
        self.sheet_combo.bind("<<ComboboxSelected>>", self._select_sheet)
        self.column_combo.bind("<<ComboboxSelected>>", self._select_column)
        self.split_office_prefix_var = tk.BooleanVar(value=False)
        self.split_office_prefix_checkbox = ttk.Checkbutton(ribbon, text="사업소 순서 접두사 (01_~15_)",
            variable=self.split_office_prefix_var, command=self._office_prefix_changed, state="disabled")
        self.split_office_prefix_checkbox.grid(row=1, column=0, columnspan=2, sticky="w", pady=(9, 0))
        self.split_office_note = tk.StringVar(value="선택 열이 기준 15개 사업소와 모두 일치할 때 사용")
        ttk.Label(ribbon, textvariable=self.split_office_note, style="Shell.TLabel", width=1).grid(row=1, column=2, sticky="ew", pady=(9, 0))
        source = ttk.Frame(page, padding=(16, 12))
        source.grid(row=1, column=0, sticky="ew")
        source.columnconfigure(0, weight=1)
        self.source_entry, _ = self._path_row(source, self.source_var, self._browse_source, "원본 파일", readonly=True)
        table = ttk.Frame(page, padding=(16, 0, 16, 12))
        table.grid(row=2, column=0, sticky="nsew")
        self.preview_tree = self._tree(table, (("label", "분류", 180), ("count", "행 수", 90), ("filename", "출력 파일", 490)))
        self._input_widgets.extend((self.sheet_combo, self.column_combo, self.pattern_entry, self.split_office_prefix_checkbox))
        self._work_footer(page, "split", "분할 시작", self._split, self.output_var, self._browse_output)
        self.progress, self.status_var = self.split_progress, self.split_status_var

    def _build_merge(self):
        page = self._page("병합", "merge")
        actions = ttk.Frame(page, padding=(16, 10))
        actions.grid(row=0, column=0, sticky="ew")
        self.merge_list_buttons = []
        for index, (label, action) in enumerate((("파일 추가", self._add_merge_files), ("선택 제거", self._remove_merge_file),
                                                ("목록 비우기", self._clear_merge_files),
                                                ("↑", lambda: self._move_merge_file(-1)), ("↓", lambda: self._move_merge_file(1)))):
            button = ttk.Button(actions, text=label, command=action, width=3 if index > 2 else None)
            button.grid(row=0, column=index, padx=(0, 6))
            self.merge_list_buttons.append(button)
            self._input_widgets.append(button)
        actions.columnconfigure(5, weight=1)
        self.merge_count_var = tk.StringVar(value="파일 0개")
        ttk.Label(actions, textvariable=self.merge_count_var, style="Note.TLabel").grid(row=0, column=5, sticky="e")
        ttk.Label(page, text="목록 순서로 병합 · 첫 파일의 서식 사용 · 표 본문 수식은 계산값으로 저장", style="Note.TLabel").grid(row=1, column=0, sticky="w", padx=16, pady=(0, 8))
        table = ttk.Frame(page, padding=(16, 0, 16, 12))
        table.grid(row=2, column=0, sticky="nsew")
        self.merge_tree = self._tree(table, (("path", "파일 이름 · 위에서 아래 순서로 병합", 420), ("sheet", "시트", 140),
                                              ("rows", "데이터 행", 100), ("status", "상태", 120)), selectmode="extended")
        self.merge_selected_path_var = tk.StringVar()
        pathbar = ttk.Frame(table)
        pathbar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(5, 0))
        pathbar.columnconfigure(0, weight=1)
        path = ttk.Label(pathbar, textvariable=self.merge_selected_path_var, width=1, style="Note.TLabel")
        path.grid(row=0, column=0, sticky="ew")
        ttk.Button(pathbar, text="경로", command=lambda: self._show_text("선택한 파일 경로", self.merge_selected_path_var.get())).grid(row=0, column=1)
        self.merge_tree.bind("<<TreeviewSelect>>", self._merge_selection_changed)
        self.merge_tree.bind("<Delete>", lambda _: self._remove_merge_file())
        self.merge_tree.bind("<Double-1>", lambda _: self._show_text("선택한 파일 경로", self.merge_selected_path_var.get()))
        self.merge_output_var = tk.StringVar()
        self._work_footer(page, "merge", "병합 시작", self._merge, self.merge_output_var, self._browse_merge_output)

    def _build_compare(self):
        page = self._page("비교", "compare")
        self.compare_reference_var, self.compare_comparison_var, self.compare_output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.compare_by_key_var = tk.BooleanVar(value=False)
        ribbon = self._ribbon(page)
        mode = self._field(ribbon, 0, "비교 방식")
        self.compare_position_radio = ttk.Radiobutton(mode, text="셀 위치", variable=self.compare_by_key_var, value=False, command=self._compare_mode_changed)
        self.compare_key_radio = ttk.Radiobutton(mode, text="키가 같은 행", variable=self.compare_by_key_var, value=True, command=self._compare_mode_changed)
        self.compare_position_radio.grid(row=1, column=0, sticky="w")
        self.compare_key_radio.grid(row=1, column=1, sticky="w", padx=(10, 0))
        self.compare_reference_table_combo = ttk.Combobox(self._field(ribbon, 1, "기준 표"), state="disabled", width=1)
        self.compare_comparison_table_combo = ttk.Combobox(self._field(ribbon, 2, "비교 표"), state="disabled", width=1)
        for combo in (self.compare_reference_table_combo, self.compare_comparison_table_combo):
            combo.grid(row=1, column=0, sticky="ew")
            combo.bind("<<ComboboxSelected>>", self._refresh_compare_keys)
        sources = ttk.Frame(page, padding=(16, 12))
        sources.grid(row=1, column=0, sticky="ew")
        for column, (title, variable, kind) in enumerate((("기준 파일", self.compare_reference_var, "reference"), ("비교 파일", self.compare_comparison_var, "comparison"))):
            sources.columnconfigure(column, weight=1, uniform="sources")
            card = ttk.Frame(sources)
            card.grid(row=0, column=column, sticky="ew", padx=(0, 12) if column == 0 else 0)
            card.columnconfigure(0, weight=1)
            self._path_row(card, variable, lambda k=kind: self._browse_compare_input(k), title, readonly=True)
        space = ttk.Frame(page, padding=(16, 0, 16, 12))
        space.grid(row=2, column=0, sticky="nsew")
        space.columnconfigure(0, weight=1)
        space.rowconfigure(0, weight=1)
        self.compare_position_note = ttk.Label(space, text="같은 이름의 시트에서 같은 위치의 셀 값을 비교합니다.\n다른 값은 결과 파일에 노란색으로 표시합니다.",
                                                style="Note.TLabel", anchor="center", justify="center")
        self.compare_position_note.grid(row=0, column=0, sticky="nsew")
        options = self.compare_key_options = ttk.Frame(space)
        options.grid(row=0, column=0, sticky="nsew")
        options.columnconfigure(0, weight=1)
        options.rowconfigure(1, weight=1)
        tools = ttk.Frame(options)
        tools.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.compare_tables_button = ttk.Button(tools, text="표 불러오기", command=self._load_compare_tables)
        self.compare_tables_button.pack(side="left")
        ttk.Label(tools, text="비교 키 · 클릭하여 여러 열 선택", style="Note.TLabel").pack(side="left", padx=10)
        self.compare_key_count_var = tk.StringVar(value="0개 선택")
        ttk.Label(tools, textvariable=self.compare_key_count_var, style="Note.TLabel").pack(side="right")
        self.compare_key_list = tk.Listbox(options, selectmode="multiple", exportselection=False, height=1, activestyle="none",
            relief="solid", borderwidth=1, highlightthickness=0, selectbackground="#edf3fc", selectforeground="#1b60bc")
        self.compare_key_list.grid(row=1, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(options, orient="vertical", command=self.compare_key_list.yview)
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(options, orient="horizontal", command=self.compare_key_list.xview)
        horizontal.grid(row=2, column=0, sticky="ew")
        self.compare_key_list.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.compare_key_list.bind("<<ListboxSelect>>", lambda _: self._render_compare_state())
        self._input_widgets.extend((self.compare_position_radio, self.compare_key_radio, self.compare_tables_button,
                                    self.compare_reference_table_combo, self.compare_comparison_table_combo, self.compare_key_list))
        self._work_footer(page, "compare", "비교 시작", self._compare, self.compare_output_var, self._browse_compare_output)

    def _build_etc(self):
        page = self._page("시트 정리", "etc")
        self.etc_source_var, self.etc_sheet_var, self.etc_output_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.etc_remove_artifacts_var, self.etc_reset_fill_var = tk.BooleanVar(value=False), tk.BooleanVar(value=False)
        self.etc_exclude_table_headers_var, self.etc_remove_conditional_formats_var = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
        ribbon = self._ribbon(page)
        ribbon.columnconfigure(1, weight=1)
        sheet = ttk.Frame(ribbon, style="Shell.TFrame")
        sheet.grid(row=0, column=0, sticky="nw", padx=(0, 22))
        ttk.Label(sheet, text="워크시트", style="Shell.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.etc_sheet_combo = ttk.Combobox(sheet, textvariable=self.etc_sheet_var, state="disabled", width=18)
        self.etc_sheet_combo.grid(row=1, column=0, sticky="ew")
        self.etc_sheet_combo.bind("<<ComboboxSelected>>", lambda _: self._render_etc_state())
        options = ttk.Frame(ribbon, style="Shell.TFrame")
        options.grid(row=0, column=1, sticky="ew")
        self.etc_remove_artifacts_checkbox = ttk.Checkbutton(options, text="도형·메모·댓글 삭제", variable=self.etc_remove_artifacts_var, command=self._render_etc_state)
        self.etc_reset_fill_checkbox = ttk.Checkbutton(options, text="채우기 색 초기화", variable=self.etc_reset_fill_var, command=self._render_etc_state)
        self.etc_exclude_table_headers_checkbox = ttk.Checkbutton(options, text="표 머리글 제외", variable=self.etc_exclude_table_headers_var, command=self._render_etc_state)
        self.etc_remove_conditional_formats_checkbox = ttk.Checkbutton(options, text="조건부 서식 삭제", variable=self.etc_remove_conditional_formats_var, command=self._render_etc_state)
        self.etc_remove_artifacts_checkbox.grid(row=0, column=0, sticky="w", padx=(0, 18), pady=3)
        self.etc_reset_fill_checkbox.grid(row=0, column=1, sticky="w", padx=(0, 18), pady=3)
        self.etc_exclude_table_headers_checkbox.grid(row=0, column=2, sticky="w", pady=3)
        self.etc_remove_conditional_formats_checkbox.grid(row=1, column=0, sticky="w", pady=3)
        source = ttk.Frame(page, padding=(16, 12))
        source.grid(row=1, column=0, sticky="ew")
        source.columnconfigure(0, weight=1)
        self._path_row(source, self.etc_source_var, self._browse_etc_source, "원본 파일", readonly=True)
        ttk.Label(page, text="선택한 시트만 정리하고 새 파일로 저장합니다.\n\n도형 삭제에는 그림·차트·버튼이 포함됩니다.\n조건부 서식 삭제는 머리글을 포함한 시트 전체에 적용됩니다.",
                  style="Note.TLabel", justify="center", anchor="center").grid(row=2, column=0, sticky="nsew", padx=16, pady=12)
        self._input_widgets.extend((self.etc_sheet_combo, self.etc_remove_artifacts_checkbox, self.etc_reset_fill_checkbox,
                                    self.etc_exclude_table_headers_checkbox, self.etc_remove_conditional_formats_checkbox))
        self._work_footer(page, "etc", "정리 시작", self._run_etc, self.etc_output_var, self._browse_etc_output)

    def _build_error_panel(self):
        self.error_panel = ttk.Frame(self.root, padding=(16, 5))
        self.error_panel.grid(row=3, column=0, sticky="ew")
        self.error_panel.columnconfigure(0, weight=1)
        self.error_message_var = tk.StringVar()
        ttk.Label(self.error_panel, textvariable=self.error_message_var, width=1, foreground="#a3372b").grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(self.error_panel, text="오류 상세", command=self._toggle_error_detail).grid(row=0, column=1)
        ttk.Button(self.error_panel, text="오류 복사", command=self._copy_error).grid(row=0, column=2, padx=(6, 0))
        self.error_text = tk.Text(self.error_panel, state="disabled")
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
        self.etc_button.configure(state="normal" if (ready or "etc" in self._result_visible) and not self._busy else "disabled")

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

    def _render_compare_state(self) -> None:
        error = self._output_error("compare", (self.compare_reference_var.get(), self.compare_comparison_var.get()))
        ready = all(variable.get() for variable in (
            self.compare_reference_var, self.compare_comparison_var, self.compare_output_var,
        ))
        key_mode = self.compare_by_key_var.get()
        for combo in (self.compare_reference_table_combo, self.compare_comparison_table_combo):
            if key_mode:
                combo.master.grid()
            else:
                combo.master.grid_remove()
        if key_mode:
            self.compare_key_options.grid()
            self.compare_position_note.grid_remove()
        else:
            self.compare_key_options.grid_remove()
            self.compare_position_note.grid()
        self.compare_key_count_var.set(f"{len(self.compare_key_list.curselection())}개 선택")
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
        self.compare_button.configure(state="normal" if (ready or "compare" in self._result_visible) and not self._busy else "disabled")

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
        self.merge_button.configure(state="normal" if (ready or "merge" in self._result_visible) and not self._busy else "disabled")
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
            note = "다음 실행에 사용할 파일명입니다. 저장된 파일은 실행 결과에서 여세요."
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
        self._show_editor("compare")
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
