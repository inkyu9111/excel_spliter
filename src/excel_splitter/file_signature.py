import hashlib
import os
from pathlib import Path

from .errors import WorkbookValidationError
from .models import FileSignature


def capture_signature(path: Path) -> FileSignature:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        before = os.fstat(stream.fileno())
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
        after = os.fstat(stream.fileno())
        if (before.st_size, before.st_mtime_ns) != (
            after.st_size,
            after.st_mtime_ns,
        ):
            raise WorkbookValidationError("파일을 읽는 동안 변경되었습니다.")
    return FileSignature(after.st_size, after.st_mtime_ns, digest.hexdigest())


def same_signature(path: Path, expected: FileSignature | None) -> bool:
    if expected is None:
        return not path.exists()
    return path.is_file() and capture_signature(path) == expected
