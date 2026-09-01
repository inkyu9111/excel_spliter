# Comment Removal and Split Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove notes and shapes from split outputs, warn before doing so, reduce Excel row-deletion calls, and allow at most three parallel Excel workers.

**Architecture:** Preview captures one artifact-presence flag using a shared Excel compatibility helper. The plaintext master is cleaned and saved once before output copies are created. Output workers use contiguous block deletion with a clean row-wise retry, while the existing calculation/save lock bounds CPU-heavy work.

**Tech Stack:** Python 3.12, tkinter, pywin32 Excel COM automation, pytest, PyInstaller

**Spec:** `docs/superpowers/specs/2026-09-01-comment-removal-performance-design.md`

## Global Constraints

- Support `.xlsx` only and require the real Microsoft Excel COM object model at runtime.
- Never modify the DRM-protected source workbook.
- Use the exact warning text `메모 및 도형은 삭제됩니다` only when the selected sheet contains a removable artifact.
- Delete legacy notes, threaded comments, and every item in the retained sheet's `Shapes` collection from the plaintext master before output copies are made.
- Preserve source/target signature checks, atomic output publication, plaintext cleanup, progress ordering, and fatal-stop behavior.
- Worker count is exactly `0` for zero targets and otherwise `min(3, target_count)`.
- Preserve the existing calculation/save serialization lock.
- Implement bug fixes test-first and keep all existing tests passing.

---

### Task 1: Artifact detection, warning, and one-time master cleanup

**Files:**
- Create: `src/excel_splitter/excel_artifacts.py`
- Modify: `src/excel_splitter/models.py`
- Modify: `src/excel_splitter/source_session.py`
- Modify: `src/excel_splitter/excel_gateway.py`
- Modify: `src/excel_splitter/gui.py`
- Test: `tests/unit/test_excel_artifacts.py`
- Test: `tests/unit/test_source_session.py`
- Test: `tests/unit/test_gui.py`
- Test: `tests/unit/test_excel_gateway.py`

**Interfaces:**
- Produces: `has_removable_artifacts(sheet: Any) -> bool` and `delete_removable_artifacts(sheet: Any) -> None` in `excel_artifacts.py`.
- Produces: `WorkbookSnapshot.has_removable_artifacts: bool` with a default of `False` for existing constructor compatibility.
- Consumes: existing unsupported-threaded-comment error classification, moved into or delegated to `excel_artifacts.py` so preview safety checks and deletion share one rule.

- [ ] **Step 1: Write failing artifact helper tests**

Add tests proving that legacy comments, supported threaded comments, and ordinary shapes each set the flag; an unsupported threaded-comment object model is treated as absent; unexpected COM errors propagate; deletion walks collections backward and leaves zero artifacts.

- [ ] **Step 2: Run the focused tests and verify RED**

Run: `pytest -q tests/unit/test_excel_artifacts.py tests/unit/test_source_session.py tests/unit/test_gui.py tests/unit/test_excel_gateway.py`

Expected: failures because the helper module, snapshot field, warning, and deletion behavior do not exist.

- [ ] **Step 3: Implement shared artifact handling and snapshot detection**

Implement:

```python
def has_removable_artifacts(sheet: Any) -> bool:
    return legacy_comment_count(sheet) > 0 or threaded_comment_count(sheet) > 0 or int(sheet.Shapes.Count) > 0

def delete_removable_artifacts(sheet: Any) -> None:
    delete_collection_backwards(sheet.Comments)
    delete_supported_threaded_comments_backwards(sheet)
    delete_collection_backwards(sheet.Shapes)
```

Treat only `AttributeError`, COM member-not-found HRESULT `-2147352573` or `-2147352570`, and HRESULT `-2147352567` with a threaded-comment-identifying message as unsupported. Propagate every other error. Build `WorkbookSnapshot(..., has_removable_artifacts=has_removable_artifacts(sheet))` during preview.

- [ ] **Step 4: Delete artifacts once from the plaintext master**

After source `SaveAs` succeeds and the in-memory workbook identity is the master path, call `delete_removable_artifacts()` on the selected sheet when the snapshot flag is true, save the master once, then continue the existing table identity, row count, package, close, and reopen validation. Never call deletion on the source path before `SaveAs`.

- [ ] **Step 5: Add the conditional GUI warning and remove shape restoration**

Append `\n\n메모 및 도형은 삭제됩니다` to the split confirmation only when `preview.snapshot.has_removable_artifacts` is true. Remove `_shape_snapshot`, `_restore_shapes`, and all calls/tests that preserve shape geometry during per-output mutation.

- [ ] **Step 6: Run focused tests and verify GREEN**

Run: `pytest -q tests/unit/test_excel_artifacts.py tests/unit/test_source_session.py tests/unit/test_gui.py tests/unit/test_excel_gateway.py`

Expected: all focused tests pass, including a regression that a shape named `Comment 4` is never restored by name.

### Task 2: Block row deletion with clean fallback and three-worker cap

**Files:**
- Modify: `src/excel_splitter/excel_gateway.py`
- Modify: `src/excel_splitter/parallel_writer.py`
- Test: `tests/unit/test_excel_gateway.py`
- Test: `tests/unit/test_parallel_writer.py`

**Interfaces:**
- Produces: `_contiguous_descending_blocks(indexes: Iterable[int]) -> tuple[tuple[int, int], ...]`, where each pair is inclusive `(first_index, last_index)` and blocks are ordered from highest rows to lowest rows.
- Produces: a target-local compatibility signal that makes `_write_one_group` discard its temporary workbook and retry exactly once with descending `ListRow.Delete` calls.
- Consumes: cleaned master output and shape-free mutation path from Task 1.

- [ ] **Step 1: Write failing block and worker-count tests**

Cover empty indexes, unordered and duplicate indexes, multiple contiguous runs, descending execution order, a recognized compatibility failure followed by fresh-copy row-wise success, a non-compatibility failure remaining fatal, and worker counts `0→0`, `1→1`, `2→2`, `3→3`, `10→3`.

- [ ] **Step 2: Run focused tests and verify RED**

Run: `pytest -q tests/unit/test_excel_gateway.py tests/unit/test_parallel_writer.py`

Expected: failures because block deletion, clean fallback, and the three-worker policy do not exist.

- [ ] **Step 3: Implement contiguous block deletion**

Normalize indexes to unique positive integers, coalesce adjacent values, and process blocks from highest to lowest. Use the range spanning the first and last `ListRow.Range` within each block and delete with an upward shift so only table-width cells move.

- [ ] **Step 4: Implement clean compatibility fallback**

When a bulk range member is missing or Excel reports error 1004, close the target workbook without saving, delete its `g-*.xlsx` temporary copy, recopy the untouched cleaned `m.xlsx` master, reopen it, and retry once with the existing descending row-wise deletion. Do not reopen the DRM source, publish any file from the failed attempt, or retry unknown COM failures.

- [ ] **Step 5: Increase and bound parallel workers**

Implement:

```python
def worker_count(target_count: int) -> int:
    return min(3, max(0, target_count))
```

Keep worker-owned COM sessions, priority scheduling, fatal-stop locking, coordinator progress delivery, and calculation/save serialization unchanged.

- [ ] **Step 6: Run focused tests and verify GREEN**

Run: `pytest -q tests/unit/test_excel_gateway.py tests/unit/test_parallel_writer.py`

Expected: all focused tests pass and the concurrency test observes no more than three simultaneous workers.

### Task 3: Documentation, full verification, and Windows package

**Files:**
- Modify: `README.md`
- Modify: `docs/superpowers/specs/2026-09-01-excel-splitter-design.md`
- Modify: `docs/superpowers/plans/2026-09-01-comment-removal-performance-implementation.md`
- Build output: `dist/ExcelSplitter.exe`

**Interfaces:**
- Consumes: Tasks 1 and 2 implementation and tests.
- Produces: user-facing documentation and a packaged Windows executable from the verified source tree.

- [ ] **Step 1: Update documentation**

State plainly that the selected sheet is retained, notes and shapes are removed after a conditional warning, the DRM source is opened only through Excel and remains unchanged, and splitting uses up to three isolated Excel workers with serialized calculation/save.

- [ ] **Step 2: Run the complete unit suite**

Run: `pytest -q`

Expected: zero failures and zero errors.

- [ ] **Step 3: Run source-quality checks**

Run: `python -m compileall -q src tests` and `git diff --check`.

Expected: both commands exit zero.

- [ ] **Step 4: Build the EXE**

Run the repository's existing `build.ps1` using its documented Python environment.

Expected: build exits zero and creates `dist/ExcelSplitter.exe` with the pywin32 runtime files bundled.

- [ ] **Step 5: Verify the packaged entry point**

Run the existing packaged self-test and inspect the EXE path, size, and SHA-256. Do not claim real Excel COM integration from this computer because it does not have the full desktop Excel application.

- [ ] **Step 6: SIP and final code review**

Cold-read the changed behavior, audit duplicated contracts, reconcile README and design wording, then dispatch a `sol high` reviewer over the full diff. Fix every critical or important finding and rerun the covering tests.
