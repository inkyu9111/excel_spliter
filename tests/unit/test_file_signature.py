import hashlib
from pathlib import Path

import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.file_signature import capture_signature, same_signature


def test_writing_a_second_byte_changes_the_signature(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"a")
    first = capture_signature(path)

    path.write_bytes(b"ab")

    assert capture_signature(path) != first
    assert not same_signature(path, first)


def test_matching_file_has_same_signature(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"content")

    assert same_signature(path, capture_signature(path))


def test_missing_path_matches_absent_expected_signature(tmp_path: Path) -> None:
    assert same_signature(tmp_path / "missing.xlsx", None)


def test_existing_path_does_not_match_absent_expected_signature(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"")

    assert not same_signature(path, None)


def test_signature_metadata_comes_from_the_hashed_open_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    requested = tmp_path / "book.xlsx"
    opened_generation = tmp_path / "opened.xlsx"
    replacement_generation = tmp_path / "replacement.xlsx"
    opened_generation.write_bytes(b"old")
    replacement_generation.write_bytes(b"replacement")
    opened_stat = opened_generation.stat()
    replacement_stat = replacement_generation.stat()
    real_open = Path.open
    real_stat = Path.stat

    def open_generation(path: Path, *args: object, **kwargs: object):
        if path == requested:
            return real_open(opened_generation, *args, **kwargs)
        return real_open(path, *args, **kwargs)

    def stat_generation(path: Path, *args: object, **kwargs: object):
        if path == requested:
            return replacement_stat
        return real_stat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "open", open_generation)
    monkeypatch.setattr(Path, "stat", stat_generation)

    signature = capture_signature(requested)

    assert signature.size == opened_stat.st_size
    assert signature.mtime_ns == opened_stat.st_mtime_ns
    assert signature.sha256 == hashlib.sha256(b"old").hexdigest()


def test_file_change_while_hashing_is_rejected(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"a")
    real_open = Path.open
    stream = real_open(path, "rb")

    class MutatingStream:
        def __init__(self) -> None:
            self.mutated = False

        def __enter__(self) -> "MutatingStream":
            return self

        def __exit__(self, *args: object) -> None:
            stream.close()

        def fileno(self) -> int:
            return stream.fileno()

        def read(self, size: int) -> bytes:
            data = stream.read(size)
            if not self.mutated:
                with real_open(path, "ab") as writer:
                    writer.write(b"b")
                self.mutated = True
            return data

    monkeypatch.setattr(Path, "open", lambda *_args, **_kwargs: MutatingStream())

    with pytest.raises(WorkbookValidationError, match="읽는 동안 변경"):
        capture_signature(path)
