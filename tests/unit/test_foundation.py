from pathlib import Path

from excel_splitter.models import CanonicalKey, FileSignature, GroupSummary


def test_foundation_models_are_hashable() -> None:
    signature = FileSignature(size=3, mtime_ns=7, sha256="abc")
    group = GroupSummary(
        key=CanonicalKey("blank", ""),
        label="",
        count=1,
        row_indexes=(1,),
    )
    assert hash(signature)
    assert group.row_indexes == (1,)
    assert Path("x.xlsx").suffix == ".xlsx"
