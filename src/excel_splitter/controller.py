from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path

from .models import Preview, SplitResult
from .ports import ProgressCallback, SplitServicePort


@dataclass(frozen=True)
class UiState:
    source: Path | None = None
    sheets: tuple[str, ...] = ()
    sheet_name: str | None = None
    columns: tuple[str, ...] = ()
    column_name: str | None = None
    output_dir: Path | None = None
    pattern: str = "%_분할"
    preview: Preview | None = None
    busy: bool = False


class AppController:
    def __init__(self, service: SplitServicePort) -> None:
        self._service = service
        self.state = UiState()

    def select_source(self, source: Path) -> UiState:
        sheets = self._service.list_sheets(source)
        self.state = UiState(
            source=source,
            sheets=sheets,
            output_dir=source.parent,
            pattern=self.state.pattern,
        )
        return self.state

    def select_sheet(self, sheet_name: str) -> UiState:
        source = self._required(self.state.source, "원본 파일")
        table = self._service.inspect_sheet(source, sheet_name)
        self.state = replace(
            self.state,
            sheet_name=sheet_name,
            columns=table.columns,
            column_name=None,
            preview=None,
        )
        return self.state

    def select_column(self, column_name: str) -> UiState:
        self.state = replace(self.state, column_name=column_name, preview=None)
        return self.state

    def select_output_dir(self, output_dir: Path) -> UiState:
        self.state = replace(self.state, output_dir=output_dir, preview=None)
        return self.state

    def set_pattern(self, pattern: str) -> UiState:
        self.state = replace(self.state, pattern=pattern, preview=None)
        return self.state

    def create_preview(self) -> Preview:
        preview = self._service.preview(
            self._required(self.state.source, "원본 파일"),
            self._required(self.state.sheet_name, "워크시트"),
            self._required(self.state.column_name, "분류 컬럼"),
            self.state.pattern,
            self._required(self.state.output_dir, "출력 폴더"),
        )
        self.state = replace(self.state, preview=preview)
        return preview

    def execute(self, overwrite: bool, progress: ProgressCallback) -> SplitResult:
        preview = self.state.preview
        if preview is None:
            raise RuntimeError("현재 설정의 미리보기가 필요합니다.")
        return self._service.execute(preview, overwrite, progress)

    @staticmethod
    def _required(value, label: str):
        if value is None:
            raise RuntimeError(f"{label}을(를) 먼저 선택하세요.")
        return value
