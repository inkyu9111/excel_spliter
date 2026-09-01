# DRM Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** DRM 원본을 Excel에서 한 번만 열어 시트·컬럼·미리보기에 재사용하고, 비보호 master에서 독립 Excel worker 두 개로 결과를 병렬 생성한다.

**Architecture:** `SourceSession`은 전용 command thread에서 source Excel과 DRM Workbook을 독점한다. `snapshot_reader`는 bulk `Value2`와 최초 key별 `Text` 조회를 순수하게 조립하며, `parallel_writer`는 비보호 master와 target 경로만 받아 독립 COM worker 1~2개를 조율한다. GUI는 COM 객체를 받지 않고 기존 controller/service 모델을 유지한다.

**Tech Stack:** Python 3.12, Tkinter, pywin32 312, `threading`, `queue`, `zipfile`, pytest 8.4.2, PyInstaller 6.22.2, Windows Excel COM

**Spec:** `docs/superpowers/specs/2026-09-01-excel-splitter-design.md`

## Global Constraints

- 입력과 출력은 `.xlsx`만 지원한다.
- source DRM 해제는 데스크톱 Excel에서 수행하며 XML 직접 읽기를 시도하지 않는다.
- 결과는 DRM 없는 일반 `.xlsx`여야 하며 조직 정책이 이를 막으면 결과를 만들지 않는다.
- source Workbook과 모든 COM proxy는 생성 thread 밖으로 전달하지 않는다.
- output worker는 최대 2개이며 각각 별도 `DispatchEx`와 COM apartment를 소유한다.
- 원본 SHA-256·크기·mtime, 분류 key, 최초 등장 순서와 원본 ListRow 인덱스 계약을 유지한다.
- 사용자 Excel 인스턴스를 조회하거나 종료하지 않는다.
- 결과 파일별 원자적 배치와 승인 없는 덮어쓰기 방지를 유지한다.
- 코드 변경은 실패하는 테스트를 먼저 확인하는 TDD 순서로 수행한다.

## Existing Data Contracts

- `FileSignature(size: int, mtime_ns: int, sha256: str)` is frozen.
- `CellSample(row_index: int, value: Any, text: str, is_error: bool=False, error_code: int|None=None)` uses 1-based original `ListRow` indexes.
- `GroupSummary(key: CanonicalKey, label: str, count: int, row_indexes: tuple[int, ...])` preserves first-seen group order.
- `TableInfo(sheet_name: str, table_name: str, columns: tuple[str, ...], row_count: int)` is detached from COM.
- `WorkbookSnapshot(source: Path, signature: FileSignature, sheet_name: str, table_name: str, column_name: str, row_count: int, groups: tuple[GroupSummary, ...])` is frozen.
- `OutputTarget(key: CanonicalKey, label: str, path: Path, prior_signature: FileSignature|None)` has one unique final path.
- `SplitResult(succeeded: tuple[Path, ...], failed: tuple[SplitFailure, ...])` returns paths in preview target order; `SplitFailure(label: str, message: str)` is target-scoped.
- `ProgressCallback` is `Callable[[int, int, str], None]`; the coordinator calls it once with `(0, total, "")`, then once per completed success or target-scoped failure.

---

### Task 1: Persistent DRM Source Session

**Files:**
- Create: `src/excel_splitter/source_session.py`
- Create: `tests/unit/test_source_session.py`

**Interfaces:**
- Produces synchronous RPC methods: `start() -> None`, `open_source(source: Path) -> SourceHandleInfo`, `inspect_table(sheet_name: str) -> TableInfo`, `build_snapshot(sheet_name: str, column_name: str) -> WorkbookSnapshot`, `save_plain_master(run_dir: Path, snapshot: WorkbookSnapshot) -> Path`, `close_source() -> None`, `shutdown() -> None`.
- Each public method except `start` and `shutdown` enqueues a private callable and blocks on `concurrent.futures.Future.result()`; the original exception is re-raised in the caller. Requests are serialized by one queue. `start`, `close_source`, and `shutdown` are idempotent; calls after shutdown raise `RuntimeError`.
- `SourceHandleInfo` is a frozen dataclass with `source`, `signature`, and `sheets` and contains no COM object.
- `save_plain_master` receives an existing app-owned run directory, creates and owns one UUID master inside it, returns that file path on success, and deletes it before raising `SplitExecutionError` on SaveAs or verification failure. The caller owns final run-directory cleanup.

- [ ] **Step 1: Write failing thread-affinity and lifecycle tests**

Create tests that inject `session_factory` and record thread IDs. Assert one Excel session and one workbook open are reused by open, inspect, and snapshot; source change closes the old workbook; shutdown orders workbook close, Excel quit, then COM uninitialize; no COM fake is returned.

- [ ] **Step 2: Run the source session tests and confirm RED**

Run: `python -m pytest tests/unit/test_source_session.py -q -p no:cacheprovider`

Expected: import failure for `excel_splitter.source_session`.

- [ ] **Step 3: Implement the command worker and source lifecycle**

Use a non-daemon `threading.Thread`, a request `queue.Queue`, and response `Future` objects. The worker creates the Excel session inside its own thread. `call` enqueues plain arguments and returns plain dataclasses. Capture source signature before open and reject disk signature changes before inspect, snapshot, and master save. `shutdown` is idempotent and joins the worker.

- [ ] **Step 4: Implement source operations with shallow inspect**

`open_source` calls `Workbooks.Open` once with existing safe options and returns worksheet names. `inspect_table` performs protection, visibility, `ListObjects.Count`, source type, row count and column checks but does not scan below the Table. `build_snapshot` performs the full below-Table validation and calls Task 2 helpers. All operations retain local COM references only inside the worker.

- [ ] **Step 5: Implement plain master SaveAs and verification**

Save to a UUID `.xlsx` under an app-owned run directory using `FileFormat=51`, empty `Password` and `WriteResPassword`, `ReadOnlyRecommended=False`, and `AddToMru=False`. Verify ZIP integrity plus `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, and `xl/_rels/workbook.xml.rels`. Reopen with empty passwords and validate worksheet/Table/column/row identities. Delete the master on any failure.

- [ ] **Step 6: Run focused tests**

Run: `python -m pytest tests/unit/test_source_session.py -q -p no:cacheprovider`

Expected: all tests pass.

### Task 2: Bulk Snapshot Reader

**Files:**
- Create: `src/excel_splitter/snapshot_reader.py`
- Create: `tests/unit/test_snapshot_reader.py`

**Interfaces:**
- Produces: `normalize_column_values(value2: object, row_count: int) -> tuple[object, ...]`.
- Produces: `build_samples(data_body_range: object, row_count: int) -> list[CellSample]`.
- Produces: `group_bulk_samples(data_body_range: object, row_count: int) -> tuple[GroupSummary, ...]`.
- `row_count` must be positive because empty Tables are rejected before these helpers. `normalize_column_values` maps a one-row scalar to a one-item tuple and an N-row one-column matrix to N values; every other shape raises `WorkbookValidationError`. `build_samples` emits 1-based row indexes and uses the existing `_excel_error_code` semantics through an injected or local equivalent.

- [ ] **Step 1: Write failing scalar, matrix, error, order and Text-call tests**

Cover one-row scalar, multi-row `((value,), ...)`, row-count mismatch, blank/error/bool/number/text canonicalization, first-seen order and first row index. A fake range must record `Cells.Item(row, 1).Text`; assert it is called exactly once per unique canonical key, not once per row.

- [ ] **Step 2: Run focused tests and confirm RED**

Run: `python -m pytest tests/unit/test_snapshot_reader.py -q -p no:cacheprovider`

Expected: import failure for `excel_splitter.snapshot_reader`.

- [ ] **Step 3: Implement value normalization and representative Text reads**

Read `data_body_range.Value2` exactly once. Build key membership from temporary samples with empty text, read `Text` only for each key's first row, then construct final `CellSample` values and reuse `group_samples`. Reject a shape that does not match `row_count` with `WorkbookValidationError`.

- [ ] **Step 4: Run classifier and snapshot tests**

Run: `python -m pytest tests/unit/test_snapshot_reader.py tests/unit/test_classifier.py -q -p no:cacheprovider`

Expected: all tests pass.

### Task 3: Parallel Output Coordinator

**Files:**
- Create: `src/excel_splitter/parallel_writer.py`
- Create: `tests/unit/test_parallel_writer.py`

**Interfaces:**
- Produces: `worker_count(target_count: int) -> int`, returning 0 for 0, 1 for 1–2, and 2 for 3 or more.
- Produces: `write_targets(master: Path, snapshot: WorkbookSnapshot, targets: tuple[OutputTarget, ...], write_one: Callable[[Any, Path, WorkbookSnapshot, OutputTarget], Path], session_factory: Callable[[], ContextManager[Any]], progress: ProgressCallback) -> SplitResult`.
- `session_factory` yields one Excel application owned for the lifetime of that worker loop. `write_one(excel, master, snapshot, target)` returns the final target path or raises `OSError` for a target-scoped filesystem failure and `ExcelSplitterError` for a fatal Excel trust failure.
- Produces: `ParallelWriteAborted(SplitExecutionError)` carrying `partial_result` and `unstarted` target paths.
- `ParallelWriteAborted.partial_result` follows preview order. `unstarted` is `tuple[Path, ...]` in preview order and excludes in-flight targets. The coordinator waits for every worker before returning or raising.

- [ ] **Step 1: Write failing scheduling, concurrency and failure tests**

Assert worker boundaries 0/1/2; two different thread IDs and sessions for 3+ targets; maximum concurrency 2; higher delete cost scheduled first with stable ties; out-of-order finishes return preview order; coordinator progress is monotonic; one `OSError` continues; a trust error stops new dequeue but joins in-flight work; every session closes.

- [ ] **Step 2: Run focused tests and confirm RED**

Run: `python -m pytest tests/unit/test_parallel_writer.py -q -p no:cacheprovider`

Expected: import failure for `excel_splitter.parallel_writer`.

- [ ] **Step 3: Implement the bounded dynamic queue**

Use a `PriorityQueue` keyed by negative delete cost and original target index. Each worker owns one `session_factory()` context for its loop. It checks a shared stop event before every dequeue, calls `write_one`, and sends plain result events. The coordinator alone calls `progress` and reconstructs `SplitResult` in target order.

- [ ] **Step 4: Implement fatal partial-result propagation and cleanup**

Treat `OSError` as a target failure. Treat `ExcelSplitterError` other than target filesystem errors as fatal, set stop, join workers, and raise `ParallelWriteAborted` with succeeded, failed and unstarted data. Do not terminate Excel processes globally.

- [ ] **Step 5: Run focused tests**

Run: `python -m pytest tests/unit/test_parallel_writer.py -q -p no:cacheprovider`

Expected: all tests pass.

### Task 4: Integrate Service, GUI, Diagnostics and Packaging

**Files:**
- Modify: `src/excel_splitter/app.py`
- Modify: `src/excel_splitter/controller.py`
- Modify: `src/excel_splitter/excel_gateway.py`
- Modify: `src/excel_splitter/gui.py`
- Modify: `src/excel_splitter/models.py`
- Modify: `src/excel_splitter/ports.py`
- Modify: `src/excel_splitter/split_service.py`
- Modify: `tests/unit/test_controller.py`
- Modify: `tests/unit/test_excel_gateway.py`
- Modify: `tests/unit/test_gui.py`
- Modify: `tests/unit/test_split_service.py`
- Modify: `scripts/build.ps1`

**Interfaces:**
- Consumes Task 1 `SourceSession`, Task 2 `group_bulk_samples`, and Task 3 `write_targets`.
- `SplitService` receives one session-backed gateway and closes it through `shutdown`.
- GUI uses one background request path for source commands and invokes shutdown during window close.

- [ ] **Step 1: Write failing integration-level unit tests**

Assert source selection, sheet selection and preview use one open source; changing source closes the previous handle; preview uses bulk reader; execute calls plain-master creation before parallel writing; GUI close waits for safe shutdown when idle; fatal parallel result displays completed, failed and unstarted counts; timing logs contain counts and elapsed times but no cell values or labels.

- [ ] **Step 2: Run relevant tests and confirm RED**

Run: `python -m pytest tests/unit/test_controller.py tests/unit/test_excel_gateway.py tests/unit/test_gui.py tests/unit/test_split_service.py -q -p no:cacheprovider`

Expected: failures show the old per-action Excel sessions and serial writer.

- [ ] **Step 3: Wire source session and bulk snapshot behavior**

Replace per-action source `_excel_session` calls with the persistent session. Keep GUI events plain and render only on the Tk main thread. Split shallow and full Table validation. Replace per-row snapshot loop with `group_bulk_samples`.

- [x] **Step 4: Wire master creation and parallel output**

Execute source signature guard, request plain master from source worker, close source, then call `write_targets`. Move current `_write_one_group` into the callable used by output workers. Preserve target signature verification, Windows no-clobber rename and recovery backup publication, filter clearing, sheet deletion, row-count verification, shape restoration and temp cleanup.

- [x] **Step 5: Add timing instrumentation and shutdown handling**

Use `time.perf_counter` and the existing rotating logger. Log operation name, elapsed time and counts only. Do not log cell values or group labels. App shutdown must close source workbook and its Excel instance before destroying the Tk root.

- [x] **Step 6: Run all unit tests**

Run: `python -m pytest tests/unit -q -p no:cacheprovider`

Expected: all tests pass. Final verified result before packaging: `118 passed`.

- [ ] **Step 7: Rebuild and verify the EXE**

Run: `powershell -ExecutionPolicy Bypass -File scripts/build.ps1`

Expected: dependency checks, all unit tests, one-folder startup probe, one-file build, embedded `pythoncom` DLL inspection and final EXE pywin32 self-test all pass; output is `dist/ExcelSplitter.exe`.

- [ ] **Step 8: Record release limitations**

Record that the current Viewer-only development PC cannot prove the DRM provider's SaveAs policy or real Excel parallel preservation. The executable is a release candidate until it passes the spec's DRM fixture and two-worker integration gates on the target Excel PC.
