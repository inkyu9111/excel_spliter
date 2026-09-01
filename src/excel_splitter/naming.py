from pathlib import Path
import re

from .errors import WorkbookValidationError
from .models import GroupSummary, OutputTarget


_INVALID_CHARACTERS = re.compile(r'[\\/:*?"<>|\[\]]')
_RESERVED_NAMES = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"com{number}" for number in range(1, 10)),
    *(f"lpt{number}" for number in range(1, 10)),
}
_UNSUPPORTED_EXTENSIONS = (".xls", ".xlsm", ".xlsb")
_MAX_ABSOLUTE_PATH_LENGTH = 218


def build_targets(
    pattern: str,
    groups: tuple[GroupSummary, ...],
    output_dir: Path,
    source: Path,
) -> tuple[OutputTarget, ...]:
    stem_pattern = _validate_and_strip_extension(pattern)
    source_key = _absolute_key(source)
    used_names: set[str] = set()
    targets: list[OutputTarget] = []

    for group in groups:
        stem = _INVALID_CHARACTERS.sub("_", stem_pattern.replace("%", group.label))
        stem = stem.rstrip(" .")
        if stem.casefold() in _RESERVED_NAMES:
            stem = f"_{stem}"
        if not stem:
            raise WorkbookValidationError("파일명 패턴의 결과가 비어 있습니다.")

        filename = _unique_filename(stem, used_names)
        path = output_dir / filename
        absolute_path = path.resolve()
        if len(str(absolute_path)) > _MAX_ABSOLUTE_PATH_LENGTH:
            raise WorkbookValidationError(
                "전체 절대 경로는 218자를 넘을 수 없습니다. "
                "파일명 패턴 또는 출력 폴더를 수정하세요."
            )
        if _absolute_key(absolute_path) == source_key:
            raise WorkbookValidationError("결과 파일이 원본 파일을 덮어쓸 수 없습니다.")

        targets.append(OutputTarget(group.key, group.label, path, None))

    return tuple(targets)


def _validate_and_strip_extension(pattern: str) -> str:
    folded = pattern.casefold()
    if folded.endswith(".xlsx"):
        pattern = pattern[:-5]
        folded = folded[:-5]
    elif folded.endswith(_UNSUPPORTED_EXTENSIONS):
        raise WorkbookValidationError("지원하지 않는 Excel 파일 확장자입니다.")
    if "%" not in pattern:
        raise WorkbookValidationError("파일명 패턴에는 %가 하나 이상 필요합니다.")
    return pattern


def _unique_filename(stem: str, used_names: set[str]) -> str:
    number = 1
    filename = f"{stem}.xlsx"
    while filename.casefold() in used_names:
        number += 1
        filename = f"{stem} ({number}).xlsx"
    used_names.add(filename.casefold())
    return filename


def _absolute_key(path: Path) -> str:
    return str(path.resolve()).casefold()
