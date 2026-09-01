from __future__ import annotations

import threading
import time
from contextlib import contextmanager
from pathlib import Path

import pytest

from excel_splitter.errors import SplitExecutionError
from excel_splitter.models import (
    CanonicalKey,
    FileSignature,
    GroupSummary,
    OutputTarget,
    WorkbookSnapshot,
)
from excel_splitter.parallel_writer import (
    ParallelWriteAborted,
    worker_count,
    write_targets,
)


def _fixture(count: int) -> tuple[WorkbookSnapshot, tuple[OutputTarget, ...]]:
    groups = tuple(
        GroupSummary(CanonicalKey("text", str(i)), str(i), i + 1, (i + 1,))
        for i in range(count)
    )
    snapshot = WorkbookSnapshot(
        Path("source.xlsx"),
        FileSignature(1, 2, "hash"),
        "Sheet1",
        "Table1",
        "Team",
        20,
        groups,
    )
    targets = tuple(
        OutputTarget(group.key, group.label, Path(f"{group.label}.xlsx"), None)
        for group in groups
    )
    return snapshot, targets


@pytest.mark.parametrize(
    ("target_count", "expected"), [(0, 0), (1, 1), (2, 1), (3, 2), (20, 2)]
)
def test_worker_count_is_bounded(target_count: int, expected: int) -> None:
    assert worker_count(target_count) == expected


def test_parallel_writes_use_two_owned_sessions_and_restore_preview_order() -> None:
    snapshot, targets = _fixture(4)
    session_threads: set[int] = set()
    active = 0
    maximum_active = 0
    lock = threading.Lock()
    progress: list[tuple[int, int, str]] = []

    @contextmanager
    def session_factory():
        session_threads.add(threading.get_ident())
        yield object()

    def write_one(_excel, _master, _snapshot, target):
        nonlocal active, maximum_active
        with lock:
            active += 1
            maximum_active = max(maximum_active, active)
        time.sleep(0.01 * (4 - int(target.label)))
        with lock:
            active -= 1
        return target.path

    result = write_targets(
        Path("master.xlsx"),
        snapshot,
        targets,
        write_one,
        session_factory,
        lambda *args: progress.append(args),
    )

    assert len(session_threads) == 2
    assert maximum_active == 2
    assert result.succeeded == tuple(target.path for target in targets)
    assert not result.failed
    assert progress[0] == (0, 4, "")
    assert [event[0] for event in progress] == [0, 1, 2, 3, 4]


def test_target_oserror_is_recorded_and_other_targets_continue() -> None:
    snapshot, targets = _fixture(3)

    @contextmanager
    def session_factory():
        yield object()

    def write_one(_excel, _master, _snapshot, target):
        if target.label == "1":
            raise OSError("locked")
        return target.path

    result = write_targets(
        Path("master.xlsx"), snapshot, targets, write_one, session_factory, lambda *_: None
    )

    assert result.succeeded == (Path("0.xlsx"), Path("2.xlsx"))
    assert [(failure.label, failure.message) for failure in result.failed] == [
        ("1", "locked")
    ]


def test_fatal_error_stops_new_work_and_reports_partial_result() -> None:
    snapshot, targets = _fixture(5)
    release_inflight = threading.Event()
    started: list[str] = []
    lock = threading.Lock()

    @contextmanager
    def session_factory():
        yield object()

    def write_one(_excel, _master, _snapshot, target):
        with lock:
            started.append(target.label)
        if target.label == "0":
            raise SplitExecutionError("session untrusted")
        release_inflight.wait(0.2)
        return target.path

    def release() -> None:
        time.sleep(0.03)
        release_inflight.set()

    threading.Thread(target=release, daemon=True).start()
    with pytest.raises(ParallelWriteAborted) as captured:
        write_targets(
            Path("master.xlsx"),
            snapshot,
            targets,
            write_one,
            session_factory,
            lambda *_: None,
        )

    error = captured.value
    assert any(failure.label == "0" for failure in error.partial_result.failed)
    assert error.unstarted
    assert set(error.unstarted).isdisjoint(error.partial_result.succeeded)
    assert len(started) < len(targets)


def test_no_target_is_claimed_after_fatal_stop_is_published(monkeypatch) -> None:
    snapshot, targets = _fixture(4)
    fatal_published = threading.Event()
    started: list[str] = []

    class PublishedFatal(SplitExecutionError):
        def __str__(self) -> str:
            fatal_published.set()
            return super().__str__()

    @contextmanager
    def session_factory():
        yield object()

    def write_one(_excel, _master, _snapshot, target):
        started.append(target.label)
        if target.label == "0":
            raise PublishedFatal("fatal")
        fatal_published.wait(1)
        return target.path

    with pytest.raises(ParallelWriteAborted):
        write_targets(
            Path("master.xlsx"), snapshot, targets, write_one,
            session_factory, lambda *_: None,
        )

    assert set(started).issubset({"0", "1"})
