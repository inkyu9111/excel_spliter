from pathlib import Path
import re

from .errors import WorkbookValidationError
from .models import GroupSummary, OutputTarget


_INVALID_CHARACTERS = re.compile(r'[\x00-\x1f\\/:*?"<>|\[\]]')
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
OFFICE_ORDER = (
    "서울", "남서울", "인천", "경기북부", "경기", "강원", "충북", "대전세종충남",
    "전북", "광주전남", "대구", "경북", "부산울산", "경남", "제주",
)
_OFFICE_RANK = {name: index for index, name in enumerate(OFFICE_ORDER, 1)}


def office_prefix_available(groups: tuple[GroupSummary, ...]) -> bool:
    return (
        len(groups) == len(OFFICE_ORDER)
        and all(group.key.kind == "text" for group in groups)
        and {group.key.value for group in groups} == set(OFFICE_ORDER)
    )


def build_targets(
    pattern: str,
    groups: tuple[GroupSummary, ...],
    output_dir: Path,
    source: Path,
    office_prefix: bool = False,
) -> tuple[OutputTarget, ...]:
    stem_pattern = _validate_and_strip_extension(pattern)
    if office_prefix:
        if not office_prefix_available(groups):
            raise WorkbookValidationError("사업소 번호는 분류 값이 지정된 15개 사업소와 정확히 일치할 때만 사용할 수 있습니다.")
        groups = tuple(sorted(groups, key=lambda group: _OFFICE_RANK[group.key.value]))
    source_key = _absolute_key(source)
    used_names: set[str] = set()
    targets: list[OutputTarget] = []

    for group in groups:
        stem = _INVALID_CHARACTERS.sub("_", stem_pattern.replace("%", group.label))
        stem = stem.rstrip(" .")
        if stem.split(".", 1)[0].rstrip().casefold() in _RESERVED_NAMES:
            stem = f"_{stem}"
        if not stem:
            raise WorkbookValidationError("파일명 패턴의 결과가 비어 있습니다.")
        if office_prefix:
            stem = f"{_OFFICE_RANK[group.key.value]:02d}_{stem}"

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
