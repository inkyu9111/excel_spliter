from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path

from .errors import WorkbookValidationError
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
    office_prefix_available: bool = False
    office_prefix: bool = False


class AppController:
    def __init__(self, service: SplitServicePort) -> None:
        self._service = service
        self.state = UiState()

    def select_source(self, source: Path) -> UiState:
        self.state = UiState(
            source=source,
            output_dir=source.parent,
            pattern=self.state.pattern,
        )
        sheets = self._service.list_sheets(source)
        self.state = replace(self.state, sheets=sheets)
        return self.state

    def select_sheet(self, sheet_name: str) -> UiState:
        source = self._required(self.state.source, "원본 파일")
        self.state = replace(
            self.state,
            sheet_name=sheet_name,
            columns=(),
            column_name=None,
            preview=None,
            office_prefix_available=False,
            office_prefix=False,
        )
        table = self._service.inspect_sheet(source, sheet_name)
        self.state = replace(self.state, columns=table.columns)
        return self.state

    def select_column(self, column_name: str) -> UiState:
        self.state = replace(
            self.state, column_name=None, preview=None,
            office_prefix_available=False, office_prefix=False,
        )
        if column_name not in self.state.columns:
            raise WorkbookValidationError("분류 컬럼을 다시 선택하세요.")
        available = self._service.office_prefix_available(
            self._required(self.state.source, "원본 파일"),
            self._required(self.state.sheet_name, "워크시트"),
            column_name,
        )
        self.state = replace(self.state, column_name=column_name, office_prefix_available=available)
        return self.state

    def set_office_prefix(self, enabled: bool) -> UiState:
        if enabled and not self.state.office_prefix_available:
            raise WorkbookValidationError("사업소 번호는 지정된 15개 사업소가 모두 있는 분류 컬럼에서만 사용할 수 있습니다.")
        self.state = replace(self.state, office_prefix=enabled, preview=None)
        return self.state

    def select_output_dir(self, output_dir: Path | None) -> UiState:
        self.state = replace(self.state, output_dir=output_dir, preview=None)
        return self.state

    def set_pattern(self, pattern: str) -> UiState:
        self.state = replace(self.state, pattern=pattern, preview=None)
        return self.state

    def create_preview(self) -> Preview:
        selection = (
            self._required(self.state.source, "원본 파일"),
            self._required(self.state.sheet_name, "워크시트"),
            self._required(self.state.column_name, "분류 컬럼"),
        )
        self.state = replace(self.state, preview=None)
        try:
            preview = self._service.preview(
                *selection, self.state.pattern,
                self._required(self.state.output_dir, "출력 폴더"),
                office_prefix=self.state.office_prefix,
            )
        except Exception:
            self._refresh_office_prefix()
            raise
        self.state = replace(self.state, preview=preview)
        return preview

    def execute(self, overwrite: bool, progress: ProgressCallback) -> SplitResult:
        preview = self.state.preview
        if preview is None:
            raise RuntimeError("현재 설정의 사전 검사 결과가 없습니다. 작업을 다시 시작하세요.")
        try:
            return self._service.execute(preview, overwrite, progress)
        except Exception:
            self.state = replace(self.state, preview=None)
            self._refresh_office_prefix()
            raise

    def _refresh_office_prefix(self) -> None:
        if not self.state.office_prefix_available:
            return
        # Keep a valid choice after output errors; changed/unreadable source
        # data must not leave the old eligibility displayed.
        try:
            available = self._service.office_prefix_available(
                self.state.source, self.state.sheet_name, self.state.column_name,
            )
        except Exception:
            available = False
        self.state = replace(
            self.state, office_prefix_available=available,
            office_prefix=self.state.office_prefix and available,
        )

    def shutdown(self) -> None:
        self._service.shutdown()

    @staticmethod
    def _required(value, label: str):
        if value is None:
            raise RuntimeError(f"{label}을(를) 먼저 선택하세요.")
        return value
