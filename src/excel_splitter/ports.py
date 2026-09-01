from __future__ import annotations

from pathlib import Path
from typing import Callable, Protocol

from .models import OutputTarget, Preview, SplitResult, TableInfo, WorkbookSnapshot

ProgressCallback = Callable[[int, int, str], None]


class ExcelGatewayPort(Protocol):
    def list_worksheets(self, source: Path) -> tuple[str, ...]: ...
    def inspect_table(self, source: Path, sheet_name: str) -> TableInfo: ...
    def build_snapshot(
        self, source: Path, sheet_name: str, column_name: str
    ) -> WorkbookSnapshot: ...
    def write_groups(
        self,
        snapshot: WorkbookSnapshot,
        targets: tuple[OutputTarget, ...],
        progress: ProgressCallback,
    ) -> SplitResult: ...


class SplitServicePort(Protocol):
    def list_sheets(self, source: Path) -> tuple[str, ...]: ...
    def inspect_sheet(self, source: Path, sheet_name: str) -> TableInfo: ...
    def preview(
        self,
        source: Path,
        sheet_name: str,
        column_name: str,
        pattern: str,
        output_dir: Path,
    ) -> Preview: ...
    def execute(
        self, preview: Preview, overwrite: bool, progress: ProgressCallback
    ) -> SplitResult: ...
