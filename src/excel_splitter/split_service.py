from __future__ import annotations

import os
from dataclasses import replace
from pathlib import Path

from .errors import WorkbookValidationError
from .file_signature import capture_signature
from .models import Preview, SplitResult, TableInfo
from .naming import build_targets
from .ports import ExcelGatewayPort, ProgressCallback


_MAX_ABSOLUTE_PATH_LENGTH = 218


class SplitService:
    def __init__(self, gateway: ExcelGatewayPort) -> None:
        self.gateway = gateway

    def list_sheets(self, source: Path) -> tuple[str, ...]:
        _validate_source_path(source)
        return self.gateway.list_worksheets(source)

    def inspect_sheet(self, source: Path, sheet_name: str) -> TableInfo:
        return self.gateway.inspect_table(source, sheet_name)

    def preview(
        self,
        source: Path,
        sheet_name: str,
        column_name: str,
        pattern: str,
        output_dir: Path,
    ) -> Preview:
        _validate_source_path(source)
        _validate_output_dir(output_dir)
        snapshot = self.gateway.build_snapshot(source, sheet_name, column_name)
        targets = build_targets(pattern, snapshot.groups, output_dir, source)
        signed = tuple(
            replace(
                target,
                prior_signature=(
                    capture_signature(target.path)
                    if target.path.exists()
                    else None
                ),
            )
            for target in targets
        )
        return Preview(
            snapshot,
            signed,
            tuple(target.path for target in signed if target.prior_signature),
        )

    def execute(
        self,
        preview: Preview,
        overwrite: bool,
        progress: ProgressCallback,
    ) -> SplitResult:
        if capture_signature(preview.snapshot.source) != preview.snapshot.signature:
            raise WorkbookValidationError(
                "원본 파일이 미리보기 이후 변경되었습니다."
            )
        if preview.collisions and not overwrite:
            raise WorkbookValidationError("기존 파일 덮어쓰기 승인이 필요합니다.")
        return self.gateway.write_groups(preview.snapshot, preview.targets, progress)

    def shutdown(self) -> None:
        self.gateway.shutdown()


def _validate_source_path(source: Path) -> None:
    if not source.exists() or not source.is_file():
        raise WorkbookValidationError("원본 파일이 존재하지 않습니다.")
    if source.suffix.casefold() != ".xlsx":
        raise WorkbookValidationError("원본 파일은 .xlsx 형식이어야 합니다.")
    absolute = source.resolve()
    if "[" in str(absolute) or "]" in str(absolute):
        raise WorkbookValidationError("원본 경로에는 [ 또는 ]를 사용할 수 없습니다.")
    if len(str(absolute)) > _MAX_ABSOLUTE_PATH_LENGTH:
        raise WorkbookValidationError("원본 절대 경로는 218자를 넘을 수 없습니다.")


def _validate_output_dir(output_dir: Path) -> None:
    if not output_dir.exists() or not output_dir.is_dir():
        raise WorkbookValidationError("출력 폴더가 존재하지 않습니다.")
    absolute = output_dir.resolve()
    if "[" in str(absolute) or "]" in str(absolute):
        raise WorkbookValidationError("출력 폴더 경로에는 [ 또는 ]를 사용할 수 없습니다.")
    if not os.access(absolute, os.W_OK):
        raise WorkbookValidationError("출력 폴더에 쓸 수 없습니다.")
