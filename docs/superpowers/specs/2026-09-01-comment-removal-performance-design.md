# Comment Removal and Split Performance Design

## Goal

Eliminate comment-backed shape restoration failures and make splitting faster without allowing unbounded Excel CPU and memory use.

## Approved behavior

- The DRM-protected source workbook is never modified.
- Preview records whether the selected sheet contains legacy notes, threaded comments, or shapes.
- When any such artifact exists, the split confirmation includes the exact line `메모 및 도형은 삭제됩니다`.
- The unprotected master copy deletes all legacy notes, threaded comments, and every item remaining in the selected sheet's `Shapes` collection once. This deliberately includes pictures, charts, buttons, controls, and Excel-generated comment shapes. Other sheets are not cleaned because they are removed from every output. Every output is copied from that cleaned master.
- Output generation does not snapshot, free-float, look up by name, or restore shapes.
- Excluded table rows are deleted in descending contiguous blocks. If an Excel compatibility error prevents block deletion, the target temporary workbook is discarded and recreated, then that target is retried with the existing row-by-row algorithm.
- Worker count is `0` for no targets, otherwise `min(3, target_count)`. Every worker owns its COM apartment and Excel instance.
- Calculation and save serialization remains in force so three workers cannot calculate and save concurrently.
- Existing source/target signature checks, atomic publication, plaintext cleanup, progress ordering, and fatal-stop behavior remain unchanged.

## Component changes

### Artifact policy

Create a shared artifact helper module that safely handles Excel versions without the threaded-comment object model. It exposes artifact detection and deletion so preview and master preparation use the same compatibility rules.

Unsupported threaded comments mean `AttributeError`, COM member-not-found HRESULT `-2147352573` or `-2147352570`, or Excel COM HRESULT `-2147352567` whose message identifies `CommentThreaded`, `threaded comment`, or `스레드 주석`. Only these cases are treated as absent; every other exception propagates. Deletion enumerates legacy comments, supported threaded comments, and remaining shapes backward because each `Delete()` mutates its collection.

`WorkbookSnapshot` stores a Boolean artifact flag. GUI confirmation uses only that flag and does not perform COM work on the UI thread.

After `SaveAs` changes the open workbook identity and full path to the run directory's plaintext `m.xlsx`, master preparation deletes artifacts from the selected sheet, saves once, verifies table identity and row count, closes the workbook, and reopens `m.xlsx` read-only for the existing package/structure validation. The original source path is never saved. Deleting artifacts therefore costs one pass per split rather than one pass per output file.

### Row deletion and parallelism

Output workers coalesce unique excluded one-based original `ListRow` indexes into inclusive contiguous ranges and process the range with the greatest indexes first. A block is removed using the table-column-width range from the first through last `ListRow`, shifted upward; headers and the totals row are outside this range. Preview's existing below-table validation remains the guard against shifting independent content.

A missing bulk range member or Excel error 1004 is a block-delete compatibility failure. The worker closes that per-target workbook without saving, deletes its `g-*.xlsx` run-directory copy, copies the already-cleaned `m.xlsx` again, and retries that target exactly once with descending `ListRow.Delete` calls. The DRM source and master are not reopened or recreated. A retry failure and every other COM failure remain fatal, and nothing from the failed attempt is published.

One, two, or three targets use one, two, or three workers respectively. Larger jobs remain capped at three. Existing priority scheduling continues to start the highest delete-cost outputs first.

## Acceptance criteria

- A workbook containing a comment shape such as `Comment 4` no longer enters shape restoration.
- The confirmation warning is conditional and uses the approved Korean text exactly.
- Successful outputs contain no legacy notes, threaded comments, or shapes on the retained sheet.
- The source workbook remains byte-for-byte unchanged.
- Output table rows and group membership remain correct in block and fallback modes.
- Two outputs run with two workers; three or more run with three workers; no run starts more than three.
- Any execution error resets the GUI progress bar through the existing error path.
- The full unit suite and packaged EXE build complete successfully.
