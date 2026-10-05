from pathlib import Path

import pytest

from excel_splitter.errors import WorkbookValidationError
from excel_splitter.models import CanonicalKey, GroupSummary
from excel_splitter.naming import OFFICE_ORDER, build_targets, office_prefix_available


def group(label: str, key: str = "x") -> GroupSummary:
    return GroupSummary(CanonicalKey("text", key), label, 1, (1,))


def test_blank_replaces_percent_with_empty_string(tmp_path: Path) -> None:
    blank = GroupSummary(CanonicalKey("blank", ""), "", 1, (1,))

    target = build_targets(
        "%_가나다", (blank,), tmp_path, tmp_path / "source.xlsx"
    )[0]

    assert target.path.name == "_가나다.xlsx"


def test_all_placeholders_are_replaced_and_terminal_xlsx_is_optional(
    tmp_path: Path,
) -> None:
    target = build_targets(
        "%-%-완료.XLSX", (group("서울"),), tmp_path, tmp_path / "source.xlsx"
    )[0]

    assert target.path.name == "서울-서울-완료.xlsx"


def test_invalid_filename_characters_are_sanitized(tmp_path: Path) -> None:
    target = build_targets(
        "%", (group('a\\b/c:d*e?f"g<h>i|j[k]'),), tmp_path,
        tmp_path / "source.xlsx",
    )[0]

    assert target.path.name == "a_b_c_d_e_f_g_h_i_j_k_.xlsx"


def test_sanitized_collisions_get_suffixes(tmp_path: Path) -> None:
    targets = build_targets(
        "%",
        (group("서울/경기", "a"), group("서울:경기", "b")),
        tmp_path,
        tmp_path / "source.xlsx",
    )

    assert [item.path.name for item in targets] == [
        "서울_경기.xlsx",
        "서울_경기 (2).xlsx",
    ]


def test_collisions_are_case_insensitive(tmp_path: Path) -> None:
    targets = build_targets(
        "%", (group("Report", "a"), group("report", "b")), tmp_path,
        tmp_path / "source.xlsx",
    )

    assert [item.path.name for item in targets] == [
        "Report.xlsx",
        "report (2).xlsx",
    ]


@pytest.mark.parametrize("label", ["CON", "com1", "Lpt9", "CON.report", "NUL.part"])
def test_windows_device_names_are_prefixed(label: str, tmp_path: Path) -> None:
    target = build_targets(
        "%", (group(label),), tmp_path, tmp_path / "source.xlsx"
    )[0]

    assert target.path.name == f"_{label}.xlsx"


@pytest.mark.parametrize("pattern", ["report", "%.xls", "%.xlsm", "%.xlsb"])
def test_invalid_patterns_are_rejected(pattern: str, tmp_path: Path) -> None:
    with pytest.raises(WorkbookValidationError):
        build_targets(pattern, (group("x"),), tmp_path, tmp_path / "source.xlsx")


def test_empty_name_after_cleanup_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(WorkbookValidationError):
        build_targets("%...", (group(""),), tmp_path, tmp_path / "source.xlsx")


def test_source_overwrite_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(WorkbookValidationError):
        build_targets(
            "source%", (group("", "x"),), tmp_path, tmp_path / "source.xlsx"
        )


def test_absolute_paths_over_218_characters_are_rejected(tmp_path: Path) -> None:
    output_dir = tmp_path / ("d" * 100)

    with pytest.raises(WorkbookValidationError, match="218"):
        build_targets(
            "%", (group("f" * 120),), output_dir, tmp_path / "source.xlsx"
        )


@pytest.mark.parametrize("change", ["missing", "extra", "space", "suffix", "blank", "error", "duplicate"])
def test_office_prefix_requires_exact_raw_text_set(tmp_path: Path, change: str) -> None:
    groups = tuple(group(name, name) for name in OFFICE_ORDER)
    if change == "missing":
        groups = groups[:-1]
    elif change == "extra":
        groups += (group("추가", "추가"),)
    elif change == "duplicate":
        groups = groups[:-1] + groups[:1]
    else:
        key = {
            "space": CanonicalKey("text", "서울 "), "suffix": CanonicalKey("text", "서울본부"),
            "blank": CanonicalKey("blank", ""), "error": CanonicalKey("error", "서울"),
        }[change]
        groups = (GroupSummary(key, "서울", 1, (1,)),) + groups[1:]
    assert not office_prefix_available(groups)
    with pytest.raises(WorkbookValidationError, match="15개"):
        build_targets("보고_%", groups, tmp_path, tmp_path / "source.xlsx", office_prefix=True)


def test_office_prefix_uses_fixed_rank_and_preserves_existing_display_pattern(tmp_path: Path) -> None:
    groups = tuple(GroupSummary(CanonicalKey("text", name), name, 2, (1, 2)) for name in reversed(OFFICE_ORDER))
    groups = groups[:-1] + (GroupSummary(CanonicalKey("text", "서울"), "서울 표시", 2, (1, 2)),)
    assert office_prefix_available(groups)
    plain = build_targets("보고_%", groups, tmp_path, tmp_path / "source.xlsx")
    prefixed = build_targets("보고_%", groups, tmp_path, tmp_path / "source.xlsx", office_prefix=True)
    assert plain[0].path.name == "보고_제주.xlsx"
    assert [target.path.name for target in prefixed[:5]] == [
        "01_보고_서울 표시.xlsx", "02_보고_남서울.xlsx", "03_보고_인천.xlsx", "04_보고_경기북부.xlsx", "05_보고_경기.xlsx",
    ]
    assert prefixed[-1].path.name == "15_보고_제주.xlsx"
