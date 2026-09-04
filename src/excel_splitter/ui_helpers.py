"""Read-only checks and short messages shared by the four operation screens."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
import re

from .errors import ExcelSplitterError, ExcelUnavailableError, WorkbookValidationError
from .merge_service import _same_file
from .naming import _INVALID_CHARACTERS, _unique_filename
from .split_service import _validate_output_dir


def validate_output_path(
    value: str, sources: Iterable[Path | str] = (), *, directory: bool = False,
    allow_existing: bool = False,
) -> str:
    """Return an inline problem; an absent split directory may be created after consent.

    This never creates or reserves a path. Services still validate immediately
    before saving, since a path can change after this UI check.
    """
    if not value.strip():
        return "출력 폴더를 입력하세요." if directory else "저장할 .xlsx 파일 경로를 입력하세요."
    try:
        path = Path(value).absolute()
        for part in path.parts[1:]:
            if part in (".", ".."):
                continue
            if (Path(part).is_reserved() or _INVALID_CHARACTERS.search(part)
                    or any(ord(char) < 32 for char in part) or part.rstrip(" .") != part):
                return "경로에 사용할 수 없는 이름이나 문자가 있습니다."
        if len(str(path)) > 218:
            return "Excel 경로는 218자를 넘을 수 없습니다. 더 짧은 경로를 입력하세요."
        if directory:
            if path.exists() and not path.is_dir():
                return "출력 경로가 폴더가 아닙니다. 다른 폴더를 선택하세요."
            ancestor = path
            while not ancestor.exists() and ancestor.parent != ancestor:
                ancestor = ancestor.parent
            _validate_output_dir(ancestor)
            return ""
        if path.suffix.casefold() != ".xlsx":
            return "결과 파일 이름은 .xlsx로 끝나야 합니다."
        for source in sources:
            if str(source) and _same_file(Path(source).resolve(), path.resolve()):
                return "원본과 다른 결과 파일 이름을 지정하세요."
        _validate_output_dir(path.parent)
        if len(str(path.parent / ".ec-00000000.xlsx")) > 218:
            return "Excel 임시 파일 경로가 너무 깁니다. 더 짧은 폴더를 선택하세요."
        if path.is_symlink():
            return "결과 경로가 바로 가기 링크입니다. 새 파일 이름을 지정하세요."
        if path.exists():
            if not path.is_file():
                return "결과 경로가 파일이 아닙니다. 새 파일 이름을 지정하세요."
            if not allow_existing:
                return "결과 파일이 이미 존재합니다. 새 파일명 추천을 눌러 주세요."
    except WorkbookValidationError as exc:
        return str(exc)
    except (OSError, ValueError):
        return "경로를 확인할 수 없습니다. 폴더 위치와 접근 권한을 확인하세요."
    return ""


def suggest_output_path(path: Path | str) -> Path:
    """Suggest an available filename without creating it or nested '(2) (2)' suffixes."""
    path = Path(path)
    names = {item.name.casefold() for item in path.parent.iterdir()} if path.parent.is_dir() else set()
    stem = _INVALID_CHARACTERS.sub("_", path.stem).rstrip(" .") or "결과"
    stem = "".join("_" if ord(char) < 32 else char for char in stem)
    if Path(stem + ".xlsx").is_reserved():
        stem = "_" + stem
    if path.suffix.casefold() == ".xlsx" and path.name.casefold() not in names and stem == path.stem:
        return path
    stem = re.sub(r" \(\d+\)$", "", stem)
    return path.with_name(_unique_filename(stem or "결과", names))


def format_elapsed(seconds: float) -> str:
    minutes, second = divmod(max(0, int(seconds)), 60)
    hour, minute = divmod(minutes, 60)
    return f"{hour}:{minute:02d}:{second:02d}" if hour else f"{minute:02d}:{second:02d}"


def describe_error(exc: Exception) -> tuple[str, str]:
    """Keep concrete validation context; put technical exception text in Details."""
    message = str(exc)
    if isinstance(exc, ExcelUnavailableError):
        return "Excel 연결을 준비하지 못했습니다.", "데스크톱 Excel 설치를 확인하세요. 계속 실패하면 오류 내용을 복사해 전달해 주세요."
    if isinstance(exc, PermissionError):
        return "파일에 접근할 수 없습니다.", "파일이 다른 프로그램에서 열려 있는지와 저장 폴더의 권한을 확인하세요."
    if isinstance(exc, WorkbookValidationError):
        summary = message if len(message) <= 300 else message[:297] + "…"
        if "중복" in message:
            return summary, "안내된 행의 키 값을 확인하거나, 행을 구분할 키 컬럼을 추가하세요."
        if "Table" in message or "Key" in message:
            return summary, "양쪽 파일의 표와 선택한 키 컬럼을 확인하세요."
        if "변경" in message:
            return summary, "파일의 최신 상태를 확인하고 미리보기 또는 표 불러오기를 다시 실행하세요."
        if "경로" in message or "파일명" in message or "폴더" in message or "이미 존재" in message:
            return summary, "저장 위치를 확인하거나 새 파일명 추천을 사용하세요."
        return summary, "입력 파일과 선택한 작업 설정을 확인한 뒤 다시 실행하세요."
    if "Excel 자동화" in message or "HRESULT" in message or "COM 오류" in message:
        return "Excel에서 작업을 완료하지 못했습니다.", "상세 내용에서 실패한 작업을 확인하세요. 원본을 Excel에서 열 수 있는지 확인하고, 계속 실패하면 오류 내용을 복사해 전달해 주세요."
    if isinstance(exc, ExcelSplitterError) and len(message) <= 300:
        return message, "상세 내용과 로그를 확인한 뒤 다시 실행하세요."
    return "작업 중 예상하지 못한 오류가 발생했습니다.", "상세 내용이나 오류 로그를 전달해 주시면 원인을 확인할 수 있습니다."
