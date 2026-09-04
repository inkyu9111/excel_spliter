from pathlib import Path

import pytest

from excel_splitter.errors import ExcelUnavailableError, SplitExecutionError, WorkbookValidationError
from excel_splitter.ui_helpers import describe_error, format_elapsed, suggest_output_path, validate_output_path


def test_output_validation_protects_original_and_existing_files(tmp_path: Path) -> None:
    source = tmp_path / "원본.xlsx"
    source.write_bytes(b"original")
    result = tmp_path / "결과.xlsx"
    result.write_bytes(b"previous")
    assert "원본" in validate_output_path(str(source), sources=(source,), allow_existing=True)
    assert "이미" in validate_output_path(str(result))
    assert validate_output_path(str(result), allow_existing=True) == ""
    assert validate_output_path(str(tmp_path / "새파일.xlsx"), sources=(source,)) == ""
    assert source.read_bytes() == b"original"
    assert result.read_bytes() == b"previous"


@pytest.mark.parametrize("name", ["", "   ", "결과.xls", "결과?.xlsx", "NUL.xlsx", "결과.xlsx."])
def test_invalid_output_name_is_rejected(tmp_path: Path, name: str) -> None:
    value = str(tmp_path / name) if name.strip() else name
    assert validate_output_path(value)


def test_missing_split_folder_is_valid_but_not_created(tmp_path: Path) -> None:
    folder = tmp_path / "새 폴더" / "하위"
    assert validate_output_path(str(folder), directory=True) == ""
    assert "폴더" in validate_output_path(str(folder / "결과.xlsx"))
    assert not folder.exists()
    existing_file = tmp_path / "file"
    existing_file.write_bytes(b"file")
    assert validate_output_path(str(existing_file / "new"), directory=True)
    assert validate_output_path(str(existing_file), directory=True)


def test_output_suggestion_advances_number_and_does_not_write(tmp_path: Path) -> None:
    for name in ("결과.xlsx", "결과 (2).xlsx"):
        (tmp_path / name).write_bytes(b"previous")
    assert suggest_output_path(tmp_path / "결과 (2).xlsx") == tmp_path / "결과 (3).xlsx"
    assert suggest_output_path(tmp_path / "다른결과.xlsx") == tmp_path / "다른결과.xlsx"
    assert not (tmp_path / "결과 (3).xlsx").exists()


@pytest.mark.parametrize("name", ["NUL.xlsx", "결과?.xlsx", "결과.xls"])
def test_recommended_filename_is_valid_even_after_typing_an_invalid_name(tmp_path: Path, name: str) -> None:
    suggestion = suggest_output_path(tmp_path / name)
    assert validate_output_path(str(suggestion)) == ""
    assert not suggestion.exists()


def test_elapsed_duration_does_not_wrap_after_an_hour() -> None:
    assert format_elapsed(65.9) == "01:05"
    assert format_elapsed(3661) == "1:01:01"


def test_error_summary_keeps_validation_context_and_hides_raw_com_tuple() -> None:
    duplicate = WorkbookValidationError("Key가 중복되었습니다: data.xlsx / Data, 행 2, 9")
    summary, action = describe_error(duplicate)
    assert "data.xlsx" in summary and "2, 9" in summary
    assert "키" in action
    summary, action = describe_error(SplitExecutionError("Excel 자동화에 실패했습니다: (-2147352567, '예외가 발생했습니다', ...)"))
    assert "-2147352567" not in summary
    assert "Excel" in summary and action
    summary, action = describe_error(ExcelUnavailableError("pywin32를 불러올 수 없습니다."))
    assert summary and action
