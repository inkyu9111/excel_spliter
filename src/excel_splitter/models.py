from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class CanonicalKey:
    kind: str
    value: str


@dataclass(frozen=True)
class FileSignature:
    size: int
    mtime_ns: int
    sha256: str


@dataclass(frozen=True)
class CellSample:
    row_index: int
    value: Any
    text: str
    is_error: bool = False
    error_code: int | None = None


@dataclass(frozen=True)
class GroupSummary:
    key: CanonicalKey
    label: str
    count: int
    row_indexes: tuple[int, ...]


@dataclass(frozen=True)
class TableInfo:
    sheet_name: str
    table_name: str
    columns: tuple[str, ...]
    row_count: int


@dataclass(frozen=True)
class WorkbookSnapshot:
    source: Path
    signature: FileSignature
    sheet_name: str
    table_name: str
    column_name: str
    row_count: int
    groups: tuple[GroupSummary, ...]


@dataclass(frozen=True)
class OutputTarget:
    key: CanonicalKey
    label: str
    path: Path
    prior_signature: FileSignature | None


@dataclass(frozen=True)
class Preview:
    snapshot: WorkbookSnapshot
    targets: tuple[OutputTarget, ...]
    collisions: tuple[Path, ...]


@dataclass(frozen=True)
class SplitFailure:
    label: str
    message: str


@dataclass(frozen=True)
class SplitResult:
    succeeded: tuple[Path, ...]
    failed: tuple[SplitFailure, ...]
