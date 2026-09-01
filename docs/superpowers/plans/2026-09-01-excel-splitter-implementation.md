# Excel Table Splitter Implementation Plan

> **Historical baseline:** DRM source reuse, bulk preview, parallel output, no-clobber publication, and the current build record are defined in [`2026-09-01-drm-performance-implementation.md`](2026-09-01-drm-performance-implementation.md). Where this baseline conflicts with that plan or the design spec, follow the DRM performance plan and [`2026-09-01-excel-splitter-design.md`](../specs/2026-09-01-excel-splitter-design.md).

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows `ExcelSplitter.exe` that splits one range-backed Excel Table by a selected column while retaining only the selected worksheet and preserving the approved workbook features.

**Architecture:** A Tkinter GUI calls a controller and `SplitService`; pure classifier, naming, and file-signature modules stay independent of Excel. `ExcelComGateway` owns every COM object inside one worker thread and edits file-level copies with a dedicated `DispatchEx` instance. Parallel workers edit disjoint files after a short shared-foundation task, then one integration worker connects the modules.

**Tech Stack:** Python 3.12, Tkinter, pywin32 312, pytest 8.4.2, PyInstaller 6.22.2, PowerShell, Windows Excel COM

**Spec:** `docs/superpowers/specs/2026-09-01-excel-splitter-design.md`

## Global Constraints

- Execution has a hard 60-minute wall-clock limit. At minute 60, stop, preserve passing work, and report unfinished items instead of extending scope.
- Code-writing agents use `gpt-5.6-sol` with `reasoning_effort: medium`; the code-review agent uses `gpt-5.6-sol` with `reasoning_effort: high`.
- The current PC has Excel Viewer only. Unit tests and mocked COM tests run here; preservation verification on real desktop Excel remains an explicit release gate.
- Support only 64-bit Windows 10/11, Python 3.12 for building, desktop Excel 2016+, `.xlsx`, and one visible range-backed `ListObject` on the selected worksheet.
- Output contains the selected worksheet only. Existing Table filter criteria are cleared; nonmatching `ListRows` are deleted by descending original row index.
- Blank cells and formula results `""` form one blank group; `%` is replaced with an empty string for that group's filename.
- Output paths are at most 218 characters, use `.xlsx`, sanitize `\ / : * ? " < > | [ ]`, and never overwrite the source.
- Preserve source identity with size, `mtime_ns`, and SHA-256. Existing outputs may be replaced only when their pre-execution signature still matches the approved signature.
- No cancellation, retry framework, plugin system, localization framework, telemetry, or support for linked/query Tables in this one-hour build.
- Implement only three user-facing exception families: `ExcelUnavailableError`, `WorkbookValidationError`, and `SplitExecutionError`; map unexpected exceptions to the last family at the GUI boundary.
- Parallel agents share one workspace. They must not run Git mutations and must edit only their assigned files. The root agent runs tests and commits returned work sequentially.

### Acceptance Contract Used by Every Agent

The design spec is the product source of truth; every coding and review agent must read Sections 5-9 before editing. For implementation coordination, these clauses are non-negotiable:

- Classification membership is the detached preview snapshot's original `ListRow` indexes. Never re-read a value to decide membership after row or sheet deletion. Blank is `None` or formula result `""`; dates remain `Value2` serial numbers; bool, number, text, and Excel error codes are distinct key kinds.
- Keep the selected worksheet only. Preserve its Table definition/style/total row/calculated-column definitions and its self-contained cells, formulas, formatting, validation, conditional formatting, merges, widths/heights, shapes/charts, and print settings as Excel adjusts them. Clear only Table filters and remove nonmatching Table rows. References to deliberately deleted sheets may become `#REF!`; show the fixed warning from the spec before execution.
- Reject protected workbook/sheet/Table, non-`xlSrcRange` or linked/query Tables, zero data rows, and independent content in the Table's column span from the first row below the Table through the bottom of `Worksheet.UsedRange`. This scan covers values, formulas, comments/notes, hyperlinks, merged cells, and non-`Normal` styles.
- Apply the filename pipeline in spec Section 7 exactly: validate pattern/extension, replace every `%`, sanitize invalid characters, trim terminal spaces/dots, protect Windows device names, add `.xlsx`, resolve case-insensitive collisions with ` (n)`, then enforce the 218-character absolute-path limit. Reject a missing/non-directory/non-writable output folder; do not create it implicitly.
- Atomicity is per output file, not per batch. A filesystem copy/replace failure records one group failure and continues; any COM identity/save/close uncertainty stops all remaining groups. Completed outputs are not rolled back.

---

## Execution Schedule and Agent Dispatch

| Minute | Work | Agent | Model |
| --- | --- | --- | --- |
| 00-04 | Worktree baseline and pinned dependency install | root | current model |
| 04-11 | Task 1 shared foundation | `foundation_writer` | `gpt-5.6-sol`, medium |
| 11-32 | Tasks 2, 3, and 4 in parallel | `core_writer`, `com_writer`, `gui_writer` | `gpt-5.6-sol`, medium |
| 32-44 | Task 5 integration | `integration_writer` | `gpt-5.6-sol`, medium |
| 44-49 | Task 6 full tests and build | root | current model |
| 49-55 | Task 7 focused code review | `code_reviewer` | `gpt-5.6-sol`, high |
| 55-57 | Task 8 P0/P1 fixes only | `review_fixer` | `gpt-5.6-sol`, medium |
| 57-60 | Task 9 final verification and handoff | root | current model |

Before Task 1, root records the initial worktree state, sets the verified bundled Python, and installs every pinned dependency once:

```powershell
git -c safe.directory=C:/dev/excelspliter status --short
$PythonExe = 'C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $PythonExe -m pip install pywin32==312 pytest==8.4.2 pyinstaller==6.22.2
```

Task 1 must reproduce the same exact pins in the requirements files. The network-dependent install requires approval. If it cannot finish within four minutes, stop execution and report the blocker because every TDD/build gate depends on it. Preserve the initial status output and never stage paths that were already dirty unless this plan explicitly owns the exact file. Use all four concurrency slots during minutes 11-32: root coordinates while three coding agents work. Task 3 may author against Task 1 interfaces in parallel, but root does not run or accept its full test gate until Task 2's helper modules pass. Start `integration_writer` only after Tasks 2 and 3 pass and after `gui_writer` releases `app.py`, because Task 5 modifies that file.

## Locked File Ownership

| Owner | Files |
| --- | --- |
| `foundation_writer` | `pyproject.toml`, `requirements.txt`, `requirements-dev.txt`, `.gitignore`, `src/excel_splitter/__init__.py`, `models.py`, `ports.py`, `errors.py`, `tests/unit/test_foundation.py` |
| `core_writer` | `classifier.py`, `naming.py`, `file_signature.py`, `tests/unit/test_classifier.py`, `test_naming.py`, `test_file_signature.py` |
| `com_writer` | `excel_gateway.py`, `tests/unit/test_excel_gateway.py` |
| `gui_writer` | `controller.py`, `gui.py`, `app.py`, `__main__.py`, `tests/unit/test_controller.py`, `scripts/build.ps1`, `scripts/smoke_test.ps1`, `README.md` |
| `integration_writer` | `split_service.py`, `tests/unit/test_split_service.py`, then `app.py` after `gui_writer` completes |
| `review_fixer` | Only files named in accepted P0/P1 review findings |

---

### Task 1: Shared Models, Ports, and Project Skeleton

**Timebox:** 7 minutes, sequential gate before parallel work

**Agent:** Spawn `foundation_writer` with `fork_turns: "none"`, `model: "gpt-5.6-sol"`, `reasoning_effort: "medium"`. Give it this task, the Global Constraints, and the design spec path. Instruct it not to commit.

**Files:**
- Create: `pyproject.toml`
- Create: `requirements.txt`
- Create: `requirements-dev.txt`
- Create: `.gitignore`
- Create: `src/excel_splitter/__init__.py`
- Create: `src/excel_splitter/models.py`
- Create: `src/excel_splitter/ports.py`
- Create: `src/excel_splitter/errors.py`
- Test: `tests/unit/test_foundation.py`

**Interfaces:**
- Produces: immutable dataclasses and protocols imported by every later task.
- Produces: `ExcelGatewayPort` and `SplitServicePort`; implementations must match these signatures exactly.

- [ ] **Step 1: Write the failing foundation import test**

```python
from pathlib import Path

from excel_splitter.models import CanonicalKey, FileSignature, GroupSummary


def test_foundation_models_are_hashable() -> None:
    signature = FileSignature(size=3, mtime_ns=7, sha256="abc")
    group = GroupSummary(
        key=CanonicalKey("blank", ""),
        label="",
        count=1,
        row_indexes=(1,),
    )
    assert hash(signature)
    assert group.row_indexes == (1,)
    assert Path("x.xlsx").suffix == ".xlsx"
```

- [ ] **Step 2: Run the test and confirm the missing-package failure**

Run:

```powershell
& $PythonExe -m pytest tests/unit/test_foundation.py -q
```

Expected: FAIL because `excel_splitter.models` does not exist.

- [ ] **Step 3: Create the package metadata and pinned dependencies**

`pyproject.toml` must configure a `src` layout and the `excel-splitter` console entry point:

```toml
[build-system]
requires = ["setuptools>=75", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "excel-splitter"
version = "0.1.0"
requires-python = ">=3.12,<3.13"
dependencies = ["pywin32==312"]

[project.scripts]
excel-splitter = "excel_splitter.app:main"

[tool.setuptools]
package-dir = {"" = "src"}

[tool.setuptools.packages.find]
where = ["src"]

[tool.pytest.ini_options]
pythonpath = ["src"]
testpaths = ["tests"]
```

`requirements.txt` contains `pywin32==312`; `requirements-dev.txt` contains `-r requirements.txt`, `pytest==8.4.2`, and `pyinstaller==6.22.2`.

- [ ] **Step 4: Implement the shared immutable models**

`src/excel_splitter/models.py` must expose these exact types:

```python
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
```

- [ ] **Step 5: Define protocols and the three exception families**

`src/excel_splitter/ports.py`:

```python
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
```

`src/excel_splitter/errors.py` contains:

```python
class ExcelSplitterError(Exception):
    """Base error safe to translate at the GUI boundary."""


class ExcelUnavailableError(ExcelSplitterError):
    pass


class WorkbookValidationError(ExcelSplitterError):
    pass


class SplitExecutionError(ExcelSplitterError):
    pass
```

- [ ] **Step 6: Run the foundation test**

Run:

```powershell
& $PythonExe -m pytest tests/unit/test_foundation.py -q
```

Expected: `1 passed`.

- [ ] **Step 7: Root verifies ownership and commits only Task 1 files**

```powershell
git -c safe.directory=C:/dev/excelspliter add pyproject.toml requirements*.txt .gitignore src/excel_splitter/__init__.py src/excel_splitter/models.py src/excel_splitter/ports.py src/excel_splitter/errors.py tests/unit/test_foundation.py
git -c safe.directory=C:/dev/excelspliter -c user.name=Codex -c user.email=codex@local commit -m "chore: scaffold Excel splitter interfaces"
```

---

### Task 2: Canonical Classification, Naming, and File Signatures

**Timebox:** 15 minutes, parallel with Tasks 3 and 4

**Agent:** `core_writer`, `gpt-5.6-sol`, medium, no Git mutations and no edits outside its locked files.

**Files:**
- Create: `src/excel_splitter/classifier.py`
- Create: `src/excel_splitter/naming.py`
- Create: `src/excel_splitter/file_signature.py`
- Test: `tests/unit/test_classifier.py`
- Test: `tests/unit/test_naming.py`
- Test: `tests/unit/test_file_signature.py`

**Interfaces:**
- Consumes: `CellSample`, `CanonicalKey`, `FileSignature`, `GroupSummary`, `OutputTarget` from Task 1.
- Produces: `canonicalize`, `group_samples`, `capture_signature`, `same_signature`, and `build_targets`.

- [ ] **Step 1: Write focused failing classifier tests**

```python
from excel_splitter.classifier import canonicalize, group_samples
from excel_splitter.models import CellSample, CanonicalKey


def test_blank_and_formula_empty_share_one_group() -> None:
    groups = group_samples(
        [CellSample(1, None, ""), CellSample(2, "", "")]
    )
    assert groups[0].key == CanonicalKey("blank", "")
    assert groups[0].row_indexes == (1, 2)


def test_bool_is_not_number_but_one_equals_one_point_zero() -> None:
    assert canonicalize(CellSample(1, True, "TRUE")) != canonicalize(
        CellSample(2, 1, "1")
    )
    assert canonicalize(CellSample(1, 1, "1")) == canonicalize(
        CellSample(2, 1.0, "1.0")
    )


def test_error_code_is_part_of_key() -> None:
    assert canonicalize(CellSample(1, None, "#N/A", True, 2042)) == CanonicalKey(
        "error", "2042"
    )
```

- [ ] **Step 2: Run classifier tests and confirm failure**

```powershell
& $PythonExe -m pytest tests/unit/test_classifier.py -q
```

Expected: FAIL because `classifier.py` does not exist.

- [ ] **Step 3: Implement deterministic canonical keys and ordered grouping**

```python
from collections import OrderedDict
from decimal import Decimal

from .errors import WorkbookValidationError
from .models import CanonicalKey, CellSample, GroupSummary


def canonicalize(sample: CellSample) -> CanonicalKey:
    if sample.is_error:
        if sample.error_code is None:
            raise WorkbookValidationError("Excel 오류 코드를 읽을 수 없습니다.")
        return CanonicalKey("error", str(sample.error_code))
    if sample.value is None or sample.value == "":
        return CanonicalKey("blank", "")
    if isinstance(sample.value, bool):
        return CanonicalKey("bool", "true" if sample.value else "false")
    if isinstance(sample.value, (int, float)):
        number = Decimal(str(sample.value)).normalize()
        if number == 0:
            number = Decimal(0)
        return CanonicalKey("number", format(number, "f"))
    if isinstance(sample.value, str):
        return CanonicalKey("text", sample.value)
    raise WorkbookValidationError(
        f"지원하지 않는 Excel 값 형식입니다: {type(sample.value).__name__}"
    )


def group_samples(samples: list[CellSample]) -> tuple[GroupSummary, ...]:
    grouped: OrderedDict[CanonicalKey, list[CellSample]] = OrderedDict()
    for sample in samples:
        grouped.setdefault(canonicalize(sample), []).append(sample)
    return tuple(
        GroupSummary(
            key=key,
            label=_label(rows[0], key),
            count=len(rows),
            row_indexes=tuple(row.row_index for row in rows),
        )
        for key, rows in grouped.items()
    )
```

`_label` returns `""` for blank. Otherwise use `CellSample.text` when nonempty and not solely `#`; fall back in this order: the original text value, the Excel error text, `TRUE`/`FALSE`, or a locale-independent numeric string with 15 significant digits. This label affects display and filenames only, never canonical equality.

- [ ] **Step 4: Write failing naming tests for blank, sanitization, collision, and limits**

```python
from pathlib import Path

import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.models import CanonicalKey, GroupSummary
from excel_splitter.naming import build_targets


def group(label: str, key: str = "x") -> GroupSummary:
    return GroupSummary(CanonicalKey("text", key), label, 1, (1,))


def test_blank_replaces_percent_with_empty_string(tmp_path: Path) -> None:
    blank = GroupSummary(CanonicalKey("blank", ""), "", 1, (1,))
    target = build_targets("%_가나다", (blank,), tmp_path, tmp_path / "source.xlsx")[0]
    assert target.path.name == "_가나다.xlsx"


def test_sanitized_collisions_get_suffixes(tmp_path: Path) -> None:
    targets = build_targets(
        "%", (group("서울/경기", "a"), group("서울:경기", "b")), tmp_path,
        tmp_path / "source.xlsx",
    )
    assert [item.path.name for item in targets] == ["서울_경기.xlsx", "서울_경기 (2).xlsx"]


def test_source_overwrite_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(WorkbookValidationError):
        build_targets("source%", (group("", "x"),), tmp_path, tmp_path / "source.xlsx")
```

- [ ] **Step 5: Implement `build_targets` in the spec's exact transformation order**

The function signature is:

```python
def build_targets(
    pattern: str,
    groups: tuple[GroupSummary, ...],
    output_dir: Path,
    source: Path,
) -> tuple[OutputTarget, ...]:
    ...
```

Use `casefold()` for collision and reserved-name comparison. Treat terminal `.xlsx` as optional, reject terminal `.xls`, `.xlsm`, and `.xlsb`, replace all `%`, sanitize invalid characters, prefix Windows device names with `_`, add numbered suffixes, and reject final absolute paths longer than 218 characters. Set `prior_signature=None`; `SplitService.preview` fills it in Task 5.

- [ ] **Step 6: Implement and test file signatures**

```python
import hashlib
from pathlib import Path

from .models import FileSignature


def capture_signature(path: Path) -> FileSignature:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    stat = path.stat()
    return FileSignature(stat.st_size, stat.st_mtime_ns, digest.hexdigest())


def same_signature(path: Path, expected: FileSignature | None) -> bool:
    if expected is None:
        return not path.exists()
    return path.is_file() and capture_signature(path) == expected
```

Test that writing a second byte changes the signature and that `same_signature(missing, None)` is true.

- [ ] **Step 7: Run the core tests**

```powershell
& $PythonExe -m pytest tests/unit/test_classifier.py tests/unit/test_naming.py tests/unit/test_file_signature.py -q
```

Expected: all tests pass.

- [ ] **Step 8: Root reviews the diff and commits only Task 2 files**

```powershell
git -c safe.directory=C:/dev/excelspliter add src/excel_splitter/classifier.py src/excel_splitter/naming.py src/excel_splitter/file_signature.py tests/unit/test_classifier.py tests/unit/test_naming.py tests/unit/test_file_signature.py
git -c safe.directory=C:/dev/excelspliter -c user.name=Codex -c user.email=codex@local commit -m "feat: add classification and output naming"
```

---

### Task 3: Excel COM Gateway

**Timebox:** 21 minutes, parallel with Tasks 2 and 4

**Agent:** `com_writer`, `gpt-5.6-sol`, medium, no Git mutations and no edits outside its locked files.

**Files:**
- Create: `src/excel_splitter/excel_gateway.py`
- Test: `tests/unit/test_excel_gateway.py`

**Interfaces:**
- Consumes: Task 1 models and `ExcelGatewayPort`; Task 2 `capture_signature`, `group_samples`, and `same_signature`.
- Produces: `ExcelComGateway` matching all four `ExcelGatewayPort` methods.

- [ ] **Step 1: Write failing tests for Table count validation and descending deletion**

```python
import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.excel_gateway import _delete_rows, _single_table


class Recorder:
    def __init__(self) -> None:
        self.deleted: list[int] = []

    def __call__(self, index: int):
        owner = self

        class Row:
            def Delete(self) -> None:
                owner.deleted.append(index)

        return Row()


def test_delete_rows_uses_descending_indexes() -> None:
    rows = Recorder()
    _delete_rows(rows, (1, 4, 2))
    assert rows.deleted == [4, 2, 1]


def test_single_table_rejects_two_tables() -> None:
    tables = type("Tables", (), {"Count": 2})()
    with pytest.raises(WorkbookValidationError, match="2개 이상"):
        _single_table(tables)
```

- [ ] **Step 2: Run the gateway tests and confirm failure**

```powershell
& $PythonExe -m pytest tests/unit/test_excel_gateway.py -q
```

Expected: FAIL because `excel_gateway.py` does not exist.

- [ ] **Step 3: Implement the COM session boundary with lazy pywin32 imports**

```python
from contextlib import contextmanager

from .errors import ExcelSplitterError, ExcelUnavailableError, SplitExecutionError


@contextmanager
def _excel_session():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ExcelUnavailableError("pywin32를 불러올 수 없습니다.") from exc

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        yield excel
    except ExcelSplitterError:
        raise
    except Exception as exc:
        raise SplitExecutionError(f"Excel 자동화에 실패했습니다: {exc}") from exc
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
```

Only this narrow cleanup `except` is allowed. Do not add retry loops or broad recovery branches.

- [ ] **Step 4: Implement workbook inspection**

`list_worksheets`, `inspect_table`, and `build_snapshot` open with `UpdateLinks=0`, `ReadOnly=True`, `AddToMru=False`, `Notify=False`, and `IgnoreReadOnlyRecommended=True`, return detached Python values, then close with `SaveChanges=False`. The write path uses the same options except `ReadOnly=False`. Set `AutomationSecurity=3` while the private instance is alive, do not refresh links/queries, and restore every application setting in `_excel_session` before `Quit`.

`inspect_table` must:

```python
def _single_table(tables):
    if tables.Count == 0:
        raise WorkbookValidationError("정식 Excel Table이 없습니다.")
    if tables.Count > 1:
        raise WorkbookValidationError("이 시트에는 Table이 2개 이상 있습니다.")
    return tables.Item(1)
```

Reject hidden sheets, protected workbook structure/sheet, `SourceType != 1` (`xlSrcRange`), zero data rows, and any unsupported content below the Table column span. The bounded scan rectangle is the Table's first through last worksheet column, starting at `table.Range.Row + table.Range.Rows.Count` and ending at the bottom row of `Worksheet.UsedRange`; an empty rectangle passes. Check values, formulas, merged cells, comments/notes, hyperlinks, and non-`Normal` styles. This one-hour build does not create a generalized workbook analyzer.

- [ ] **Step 5: Implement snapshot extraction without leaking COM objects**

Capture the source signature before opening. Repeat every `inspect_table` validation, then require an exact case-sensitive match for one selected Table column name. For each Table row, read that cell's `Value2` and `Text`. Detect an Excel error with the COM variant/error value and pass `is_error=True` plus the numeric code into `CellSample`. Close without saving, capture the source signature again, and reject the preview if the signatures differ. Call Task 2 `group_samples` and return a fully detached `WorkbookSnapshot` carrying the verified signature. The stored `row_indexes` are the immutable membership oracle used for all outputs; no write step may reclassify a cell.

- [ ] **Step 6: Implement one-session group writing**

`write_groups` first creates a unique `.xlsx` master beside the targets with `shutil.copy2`. It must verify the source signature both immediately before and immediately after the copy, then verify the master's size, `mtime_ns`, and SHA-256 equal the snapshot signature. Any mismatch deletes the master and raises `SplitExecutionError` before Excel opens. It then performs this exact loop:

```python
master = _copy_to_master(snapshot.source)
try:
    with _excel_session() as excel:
        for completed, target in enumerate(targets, start=1):
            progress(completed - 1, len(targets), target.label)
            try:
                output = _write_one_group(excel, master, snapshot, target)
                succeeded.append(output)
            except OSError as exc:
                failed.append(SplitFailure(target.label, str(exc)))
            progress(completed, len(targets), target.label)
finally:
    master.unlink(missing_ok=True)
```

`_write_one_group` copies the master to a unique `.xlsx` temporary file in `target.path.parent`, opens only that copy, and repeats the selected sheet/Table/source type/column/row-count identity checks before editing. It clears filters, computes the complement of `target.row_indexes`, deletes those original indexes in descending order without reading cell values, deletes every other `Workbook.Sheets` item, verifies one sheet and the expected Table row count, restores each shape's snapshotted `(Left, Top, Width, Height)`, calls `workbook.Save()`, and closes it before placement. Only after a successful close may it verify `target.path` still matches `prior_signature` and call `os.replace(temp_path, target.path)`. Clean the temporary file in `finally`.

If workbook structure or COM identity becomes unreliable, raise `SplitExecutionError` out of the group loop so remaining groups stop. A destination filesystem copy or replace `OSError` becomes one `SplitFailure` and processing continues; an Excel open/save COM failure stops the session because its safety cannot be inferred within this build's timebox.

- [ ] **Step 7: Add mocked cleanup and group-failure tests**

Patch lazy `pythoncom` and `DispatchEx` imports through `sys.modules`; assert `Quit` and `CoUninitialize` run when workbook opening raises. Test `_delete_rows`, `_single_table`, `_remove_other_sheets`, selected-column rejection, the target-signature guard, and the copied-master signature guard as isolated helpers rather than mocking the entire Excel object model.

- [ ] **Step 8: Run gateway tests**

```powershell
& $PythonExe -m pytest tests/unit/test_classifier.py tests/unit/test_file_signature.py tests/unit/test_excel_gateway.py -q
```

Root runs this gate only after Task 2 is green. Expected: all tests pass without desktop Excel installed.

- [ ] **Step 9: Root reviews the diff and commits only Task 3 files**

```powershell
git -c safe.directory=C:/dev/excelspliter add src/excel_splitter/excel_gateway.py tests/unit/test_excel_gateway.py
git -c safe.directory=C:/dev/excelspliter -c user.name=Codex -c user.email=codex@local commit -m "feat: add Excel COM gateway"
```

---

### Task 4: Tkinter GUI, Controller, Packaging, and User Guide

**Timebox:** 21 minutes, parallel with Tasks 2 and 3

**Agent:** `gui_writer`, `gpt-5.6-sol`, medium, no Git mutations and no edits outside its locked files.

**Files:**
- Create: `src/excel_splitter/controller.py`
- Create: `src/excel_splitter/gui.py`
- Create: `src/excel_splitter/app.py`
- Create: `src/excel_splitter/__main__.py`
- Test: `tests/unit/test_controller.py`
- Create: `scripts/build.ps1`
- Create: `scripts/smoke_test.ps1`
- Create: `README.md`

**Interfaces:**
- Consumes: Task 1 `SplitServicePort`, `Preview`, `SplitResult`, and `TableInfo` only. It must not import `ExcelComGateway` directly.
- Produces: `AppController`, `ExcelSplitterGui`, and `main()`; Task 5 changes only the composition inside `app.py`.

- [ ] **Step 1: Write a failing controller state-reset test with a fake service**

```python
from pathlib import Path

from excel_splitter.controller import AppController
from excel_splitter.models import TableInfo


class FakeService:
    def list_sheets(self, source: Path) -> tuple[str, ...]:
        return ("분류표", "참조")

    def inspect_sheet(self, source: Path, sheet_name: str) -> TableInfo:
        return TableInfo(sheet_name, "Table1", ("구분", "금액"), 2)


def test_selecting_new_file_resets_downstream_state() -> None:
    controller = AppController(FakeService())
    controller.select_source(Path("a.xlsx"))
    controller.select_sheet("분류표")
    controller.select_source(Path("b.xlsx"))
    assert controller.state.sheet_name is None
    assert controller.state.columns == ()
    assert controller.state.preview is None
```

- [ ] **Step 2: Run the controller test and confirm failure**

```powershell
& $PythonExe -m pytest tests/unit/test_controller.py -q
```

Expected: FAIL because `controller.py` does not exist.

- [ ] **Step 3: Implement a small synchronous controller and immutable UI state**

```python
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
```

`AppController` exposes `select_source`, `select_sheet`, `select_column`, `select_output_dir`, `set_pattern`, `create_preview`, and `execute`. Source change resets sheet/column/preview and defaults output to the source folder; sheet change resets column/preview; column, output-folder, or pattern change invalidates preview. `execute` accepts only the currently stored preview so the GUI cannot run stale filenames.

- [ ] **Step 4: Build the single-window Tkinter view**

Use `ttk.Entry`, `ttk.Combobox`, `ttk.Treeview`, `ttk.Progressbar`, and buttons for source/output folder browsing, preview, and split. Show unique label/count/output filename rows in one `Treeview`. Render blank as `∅ (빈 셀)` while keeping the real label empty.

Do COM/service calls in one daemon worker thread and return events through `queue.Queue`. Poll with `root.after(75, poll_queue)` and handle exactly `("ok", payload)`, `("error", exception)`, and `("progress", completed, total, label)`. The service progress callback only enqueues `progress`; Tk widgets are touched only by `poll_queue`. Disable every input and replace the close protocol with a no-op while busy; restore both on success or failure.

```python
def _start_worker(self, action: Callable[[], object]) -> None:
    self._set_busy(True)

    def run() -> None:
        try:
            self.events.put(("ok", action()))
        except Exception as exc:
            self.events.put(("error", exc))

    threading.Thread(target=run, daemon=True).start()
```

Preview completion renders each group label/count/target filename. Clicking Split first shows the fixed warning that deleting other sheets may break their dependent formulas/names/charts/validations/connections. If collisions exist, the same confirmation includes the complete collision list; cancel leaves the preview intact, approval calls `execute(..., overwrite=True)`, and no-collision approval uses `overwrite=False`. Completion shows succeeded paths and each group failure in one summary dialog; a batch-level exception shows an error and no success claim. Any upstream field edit removes the displayed preview and disables Split until preview is rebuilt.

At the GUI boundary, show `ExcelSplitterError` messages directly and map unexpected exceptions to `"처리 중 예상하지 못한 오류가 발생했습니다."`; detailed exceptions go to `RotatingFileHandler` at `%LOCALAPPDATA%\ExcelSplitter\logs\excel-splitter.log`, 1MB, three backups.

- [ ] **Step 5: Implement the temporary composition root**

`app.py` must define `main(service: SplitServicePort | None = None)`. When `service` is `None`, raise a clear temporary `RuntimeError("SplitService wiring is pending")`; Task 5 replaces this line after all parallel tasks finish. `__main__.py` calls `main()`.

- [ ] **Step 6: Write the build and smoke scripts**

`scripts/build.ps1` accepts `-PythonExe`, verifies the already installed pinned dependencies, runs unit tests, builds one-folder first with this exact command:

```powershell
& $PythonExe -m PyInstaller --noconfirm --clean --onedir --windowed `
    --name ExcelSplitter --paths src src/excel_splitter/__main__.py
```

After checking `dist\ExcelSplitter\ExcelSplitter.exe` starts, remove only that generated `dist\ExcelSplitter` directory and `build\ExcelSplitter`, then build the final one-file executable:

```powershell
& $PythonExe -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name ExcelSplitter --paths src src/excel_splitter/__main__.py
```

Fail immediately on a nonzero exit code. `scripts/smoke_test.ps1` accepts `-ExePath` and `-WorkbookPath`; it refuses to run without desktop Excel and prints a deterministic fixture oracle. The fixture must have visible sheet `분류표`, range Table `Table1`, classification column `구분`, and rows `A`, `A`, blank, plus a second `참조` sheet. On `분류표`, include a total row, calculated column, formula, conditional format, validation, merge outside the Table columns, nondefault width/height, shape, chart, and print setting. Expected `%_결과` outputs are `A_결과.xlsx` with two data rows and `_결과.xlsx` with one, each containing only `분류표`, while the source SHA-256 is unchanged. The manual checklist records preservation-contract items, the fixed deleted-sheet-reference warning, and absence of orphan Excel processes.

- [ ] **Step 7: Write a concise README**

Document the `.xlsx`/single-Table requirements, `%` filename rule, blank filename behavior, Excel installation requirement, build command, unit-test command, output-folder requirements, partial-success behavior, fixed deleted-sheet-reference warning, and unsupported linked-Table/Table-below-content cases. Link the design spec for the complete preservation contract.

- [ ] **Step 8: Run controller tests**

```powershell
& $PythonExe -m pytest tests/unit/test_controller.py -q
```

Expected: all tests pass without creating a Tk root window.

- [ ] **Step 9: Root reviews the diff and commits only Task 4 files**

```powershell
git -c safe.directory=C:/dev/excelspliter add src/excel_splitter/controller.py src/excel_splitter/gui.py src/excel_splitter/app.py src/excel_splitter/__main__.py tests/unit/test_controller.py scripts/build.ps1 scripts/smoke_test.ps1 README.md
git -c safe.directory=C:/dev/excelspliter -c user.name=Codex -c user.email=codex@local commit -m "feat: add desktop workflow and packaging"
```

---

### Task 5: Split Service and Application Integration

**Timebox:** 12 minutes, starts after Tasks 2-4 are committed

**Agent:** `integration_writer`, `gpt-5.6-sol`, medium, no Git mutations.

**Files:**
- Create: `src/excel_splitter/split_service.py`
- Test: `tests/unit/test_split_service.py`
- Modify: `src/excel_splitter/app.py`

**Interfaces:**
- Consumes: all exact interfaces from Tasks 1-4.
- Produces: concrete `SplitService` and final `main()` dependency wiring.

- [ ] **Step 1: Write failing service tests with a fake gateway**

```python
from pathlib import Path

import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.file_signature import capture_signature
from excel_splitter.models import (
    CanonicalKey, GroupSummary, SplitResult, WorkbookSnapshot,
)
from excel_splitter.split_service import SplitService


class FakeGateway:
    def __init__(self, snapshot: WorkbookSnapshot) -> None:
        self.snapshot = snapshot
        self.written = False

    def list_worksheets(self, source: Path) -> tuple[str, ...]:
        return ("분류표",)

    def build_snapshot(self, source: Path, sheet_name: str, column_name: str):
        return self.snapshot

    def write_groups(self, snapshot, targets, progress):
        self.written = True
        return SplitResult(tuple(target.path for target in targets), ())


def test_execute_requires_overwrite_approval_when_collision_exists(tmp_path: Path):
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"source")
    signature = capture_signature(source)
    group = GroupSummary(CanonicalKey("text", "A"), "A", 1, (1,))
    snapshot = WorkbookSnapshot(source, signature, "분류표", "Table1", "구분", 1, (group,))
    service = SplitService(FakeGateway(snapshot))
    (tmp_path / "A.xlsx").write_bytes(b"existing")
    preview = service.preview(source, "분류표", "구분", "%", tmp_path)
    with pytest.raises(WorkbookValidationError, match="덮어쓰기 승인"):
        service.execute(preview, overwrite=False, progress=lambda *_: None)
```

Add a second test that appends one byte to the source after preview and asserts `gateway.write_groups` is not called. Add a third test that a missing output directory is rejected before `build_targets`.

- [ ] **Step 2: Run service tests and confirm failure**

```powershell
& $PythonExe -m pytest tests/unit/test_split_service.py -q
```

Expected: FAIL because `split_service.py` does not exist.

- [ ] **Step 3: Implement the thin orchestration service**

```python
class SplitService:
    def __init__(self, gateway: ExcelGatewayPort) -> None:
        self.gateway = gateway

    def list_sheets(self, source: Path) -> tuple[str, ...]:
        _validate_source_path(source)
        return self.gateway.list_worksheets(source)

    def inspect_sheet(self, source: Path, sheet_name: str) -> TableInfo:
        return self.gateway.inspect_table(source, sheet_name)

    def preview(self, source, sheet_name, column_name, pattern, output_dir):
        _validate_source_path(source)
        _validate_output_dir(output_dir)
        snapshot = self.gateway.build_snapshot(source, sheet_name, column_name)
        targets = build_targets(pattern, snapshot.groups, output_dir, source)
        signed = tuple(
            replace(
                target,
                prior_signature=capture_signature(target.path)
                if target.path.exists() else None,
            )
            for target in targets
        )
        return Preview(snapshot, signed, tuple(t.path for t in signed if t.prior_signature))

    def execute(self, preview, overwrite, progress):
        if capture_signature(preview.snapshot.source) != preview.snapshot.signature:
            raise WorkbookValidationError("원본 파일이 미리보기 이후 변경되었습니다.")
        if preview.collisions and not overwrite:
            raise WorkbookValidationError("기존 파일 덮어쓰기 승인이 필요합니다.")
        return self.gateway.write_groups(preview.snapshot, preview.targets, progress)
```

`_validate_source_path` checks existence, `.xlsx`, `[`/`]`, and a 218-character absolute path. `_validate_output_dir` requires an existing writable directory and never creates it. Do not duplicate COM-level workbook validation here; filesystem `OSError` at actual placement remains the authoritative race-safe writeability check.

- [ ] **Step 4: Replace temporary `app.py` wiring**

```python
def main(service: SplitServicePort | None = None) -> None:
    concrete = service or SplitService(ExcelComGateway())
    root = tk.Tk()
    ExcelSplitterGui(root, AppController(concrete))
    root.mainloop()
```

- [ ] **Step 5: Run service and controller tests together**

```powershell
& $PythonExe -m pytest tests/unit/test_split_service.py tests/unit/test_controller.py -q
```

Expected: all tests pass.

- [ ] **Step 6: Root reviews the diff and commits Task 5**

```powershell
git -c safe.directory=C:/dev/excelspliter add src/excel_splitter/split_service.py src/excel_splitter/app.py tests/unit/test_split_service.py
git -c safe.directory=C:/dev/excelspliter -c user.name=Codex -c user.email=codex@local commit -m "feat: integrate split workflow"
```

---

### Task 6: Full Unit Verification and Executable Build

**Timebox:** 5 minutes

**Owner:** root

**Files:**
- Verify: all source and unit-test files
- Generate: `dist/ExcelSplitter.exe` and intermediate build outputs ignored by Git

- [ ] **Step 1: Run the complete unit suite**

```powershell
& $PythonExe -m pytest -q
```

Expected: all tests pass; no test opens desktop Excel or Tk's main loop.

- [ ] **Step 2: Compile every Python module**

```powershell
& $PythonExe -m compileall -q src
```

Expected: exit code 0.

- [ ] **Step 3: Run the one-folder and one-file build script**

```powershell
& .\scripts\build.ps1 -PythonExe $PythonExe
```

Expected: `dist\ExcelSplitter.exe` exists. If dependency installation is blocked by network approval or the build exceeds the timebox, record the exact failing command and continue to review without claiming an executable was built.

- [ ] **Step 4: Inspect Git status and commit only deliberate build-script corrections**

```powershell
git -c safe.directory=C:/dev/excelspliter status --short
git -c safe.directory=C:/dev/excelspliter diff --check
```

Do not commit `build/`, `dist/`, `.pytest_cache/`, or `__pycache__/`.

---

### Task 7: High-Reasoning Code Review

**Timebox:** 6 minutes

**Agent:** Spawn `code_reviewer` with `fork_turns: "none"`, `model: "gpt-5.6-sol"`, `reasoning_effort: "high"`. It is read-only and must not edit files.

**Review prompt:**

```text
Review the current Excel Splitter implementation against:
- docs/superpowers/specs/2026-09-01-excel-splitter-design.md
- docs/superpowers/plans/2026-09-01-excel-splitter-implementation.md

Read all source and tests. Run tests if useful. Focus on defects that can corrupt the source, misclassify rows, overwrite unapproved files, leak/kill the wrong Excel process, cross COM thread boundaries, freeze Tkinter, or make the executable fail at startup.

Return findings only, ordered P0/P1/P2, with exact file and line. Distinguish verified defects from untested risk. Do not edit files. Ignore cosmetic refactors and feature expansion because the build has a hard one-hour limit.
```

- [ ] **Step 1: Root validates every finding against code and test evidence**

Reject speculative findings that contradict the locked scope. Accept only findings with a concrete failing path or missing required behavior.

- [ ] **Step 2: Create a short fix list containing accepted P0 and P1 findings only**

P2 findings go into the final handoff as deferred items; they do not extend the one-hour build.

---

### Task 8: Fix Accepted P0/P1 Review Findings

**Timebox:** 2 minutes

**Agent:** Spawn `review_fixer` with `fork_turns: "none"`, `model: "gpt-5.6-sol"`, `reasoning_effort: "medium"`. Give it only accepted findings, exact file ownership, and the failing test command. No Git mutations.

- [ ] **Step 1: Add one focused failing regression test per accepted finding**

Each regression test must reproduce the concrete review path. Do not broaden the test matrix.

- [ ] **Step 2: Run only those tests and confirm they fail for the reviewed reason**

```powershell
$AcceptedTestFiles = @('tests/unit/test_excel_gateway.py') # replace this list with the accepted findings' exact test files
& $PythonExe -m pytest $AcceptedTestFiles -q
```

The root sets the array to only the test files recorded in the accepted-finding fix list. Do not run the full suite inside this two-minute fix task.

- [ ] **Step 3: Apply the smallest fix and rerun the same focused tests**

Expected: focused tests pass.

- [ ] **Step 4: Root reviews and commits fixes**

```powershell
$AcceptedFiles = @('src/excel_splitter/excel_gateway.py', 'tests/unit/test_excel_gateway.py') # replace with exact reviewed paths
git -c safe.directory=C:/dev/excelspliter add -- $AcceptedFiles
git -c safe.directory=C:/dev/excelspliter -c user.name=Codex -c user.email=codex@local commit -m "fix: address Excel splitter review findings"
```

Set the array to explicit reviewed file paths and compare them with the initial dirty-worktree baseline before staging. Never use a directory-wide add in this task.

If no P0/P1 findings survive validation, skip this task and preserve the remaining minutes for final verification.

---

### Task 9: Final Verification and Handoff

**Timebox:** final 3 minutes; stop at the hard 60-minute boundary

**Owner:** root

- [ ] **Step 1: Run fresh verification**

```powershell
& $PythonExe -m pytest -q
& $PythonExe -m compileall -q src
git -c safe.directory=C:/dev/excelspliter diff --check
git -c safe.directory=C:/dev/excelspliter status --short
```

Claim passing status only from this fresh output. If `dist\ExcelSplitter.exe` exists, record its size and SHA-256:

```powershell
Get-Item .\dist\ExcelSplitter.exe | Select-Object FullName,Length
Get-FileHash .\dist\ExcelSplitter.exe -Algorithm SHA256
```

- [ ] **Step 2: Record the real-Excel release gate**

State explicitly that the current PC cannot verify COM preservation. Provide this command for the target Excel PC:

```powershell
& .\scripts\smoke_test.ps1 -ExePath .\dist\ExcelSplitter.exe -WorkbookPath C:\ExcelSplitterSmoke\preservation-smoke.xlsx
```

The release gate checks selected-sheet-only output, group row membership, blank output naming, source hash preservation, format/chart/shape preservation, and absence of orphan Excel processes.

- [ ] **Step 3: Hand off outcome without extending scope**

Report:

- implemented modules and commit IDs;
- exact unit-test count and result;
- executable path/hash or exact build blocker;
- P0/P1 fixes and deferred P2 findings;
- real-Excel smoke test status as unrun, passed, or failed;
- any work stopped at minute 60.
