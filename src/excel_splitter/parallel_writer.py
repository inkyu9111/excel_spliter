from __future__ import annotations

import queue
import threading
from collections.abc import Callable
from contextlib import AbstractContextManager
from pathlib import Path
from typing import Any

from .errors import ExcelSplitterError, SplitExecutionError
from .models import OutputTarget, SplitFailure, SplitResult, WorkbookSnapshot
from .ports import ProgressCallback


WriteOne = Callable[[Any, Path, WorkbookSnapshot, OutputTarget], Path]
SessionFactory = Callable[[], AbstractContextManager[Any]]


class ParallelWriteAborted(SplitExecutionError):
    def __init__(
        self,
        message: str,
        partial_result: SplitResult,
        unstarted: tuple[Path, ...],
    ) -> None:
        super().__init__(message)
        self.partial_result = partial_result
        self.unstarted = unstarted


def worker_count(target_count: int) -> int:
    if target_count <= 0:
        return 0
    if target_count <= 2:
        return 1
    return 2


def write_targets(
    master: Path,
    snapshot: WorkbookSnapshot,
    targets: tuple[OutputTarget, ...],
    write_one: WriteOne,
    session_factory: SessionFactory,
    progress: ProgressCallback,
) -> SplitResult:
    total = len(targets)
    progress(0, total, "")
    count = worker_count(total)
    if count == 0:
        return SplitResult((), ())

    tasks: queue.PriorityQueue[tuple[int, int, OutputTarget]] = queue.PriorityQueue()
    group_counts = {group.key: group.count for group in snapshot.groups}
    for index, target in enumerate(targets):
        delete_cost = snapshot.row_count - group_counts[target.key]
        tasks.put((-delete_cost, index, target))

    events: queue.Queue[tuple[str, int | None, object]] = queue.Queue()
    stop = threading.Event()
    started: set[int] = set()
    scheduling_lock = threading.Lock()

    def claim() -> tuple[int, OutputTarget] | None:
        with scheduling_lock:
            if stop.is_set():
                return None
            try:
                _cost, index, target = tasks.get_nowait()
            except queue.Empty:
                return None
            started.add(index)
            return index, target

    def publish_stop() -> None:
        with scheduling_lock:
            stop.set()

    def worker() -> None:
        try:
            with session_factory() as excel:
                while True:
                    claimed = claim()
                    if claimed is None:
                        break
                    index, target = claimed
                    try:
                        output = write_one(excel, master, snapshot, target)
                    except OSError as exc:
                        events.put(("failure", index, SplitFailure(target.label, str(exc))))
                    except ExcelSplitterError as exc:
                        publish_stop()
                        events.put(("fatal", index, exc))
                        break
                    except Exception as exc:
                        publish_stop()
                        events.put(
                            (
                                "fatal",
                                index,
                                SplitExecutionError(f"Excel 병렬 작업에 실패했습니다: {exc}"),
                            )
                        )
                        break
                    else:
                        events.put(("success", index, output))
        except Exception as exc:
            publish_stop()
            error = (
                exc
                if isinstance(exc, ExcelSplitterError)
                else SplitExecutionError(f"Excel worker를 시작하지 못했습니다: {exc}")
            )
            events.put(("fatal", None, error))
        finally:
            events.put(("worker_done", None, None))

    threads = [threading.Thread(target=worker, daemon=False) for _ in range(count)]
    for thread in threads:
        thread.start()

    successes: dict[int, Path] = {}
    failures: dict[int, SplitFailure] = {}
    fatal_errors: list[ExcelSplitterError] = []
    completed = 0
    workers_done = 0
    while workers_done < count:
        kind, index, payload = events.get()
        if kind == "worker_done":
            workers_done += 1
            continue
        if kind == "success" and index is not None:
            successes[index] = payload  # type: ignore[assignment]
        elif kind == "failure" and index is not None:
            failures[index] = payload  # type: ignore[assignment]
        elif kind == "fatal":
            fatal = payload
            if isinstance(fatal, ExcelSplitterError):
                fatal_errors.append(fatal)
            if index is not None:
                failures[index] = SplitFailure(targets[index].label, str(fatal))
        if index is not None:
            completed += 1
            progress(completed, total, targets[index].label)

    for thread in threads:
        thread.join()

    result = SplitResult(
        tuple(successes[index] for index in range(total) if index in successes),
        tuple(failures[index] for index in range(total) if index in failures),
    )
    if fatal_errors:
        with scheduling_lock:
            unstarted = tuple(
                target.path for index, target in enumerate(targets) if index not in started
            )
        raise ParallelWriteAborted(str(fatal_errors[0]), result, unstarted)
    return result
