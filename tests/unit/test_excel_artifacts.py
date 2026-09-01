from __future__ import annotations

from types import SimpleNamespace

import pytest

from excel_splitter.excel_artifacts import (
    delete_removable_artifacts,
    has_removable_artifacts,
)


class MutableCollection:
    def __init__(self, names: tuple[str, ...]) -> None:
        self.names = list(names)
        self.deleted: list[str] = []

    @property
    def Count(self) -> int:
        return len(self.names)

    def Item(self, index: int):
        name = self.names[index - 1]

        def delete() -> None:
            self.deleted.append(name)
            self.names.pop(index - 1)

        return SimpleNamespace(Delete=delete)


def sheet_with(
    *, comments: int = 0, threaded_comments: int = 0, shapes: int = 0
) -> SimpleNamespace:
    return SimpleNamespace(
        Comments=SimpleNamespace(Count=comments),
        CommentsThreaded=SimpleNamespace(Count=threaded_comments),
        Shapes=SimpleNamespace(Count=shapes),
    )


@pytest.mark.parametrize(
    "sheet",
    (
        sheet_with(comments=1),
        sheet_with(threaded_comments=1),
        sheet_with(shapes=1),
    ),
)
def test_each_supported_artifact_marks_the_sheet_for_cleanup(sheet: object) -> None:
    assert has_removable_artifacts(sheet) is True


@pytest.mark.parametrize(
    ("hresult", "message"),
    (
        (-2147352573, "member not found"),
        (-2147352570, "unknown name"),
        (-2147352567, "CommentThreaded is unavailable"),
        (-2147352567, "threaded comment is unavailable"),
        (-2147352567, "스레드 주석을 지원하지 않습니다"),
    ),
)
def test_unsupported_threaded_comment_models_are_treated_as_absent(
    hresult: int, message: str
) -> None:
    com_error = type("com_error", (Exception,), {})

    class LegacySheet:
        Comments = SimpleNamespace(Count=0)
        Shapes = SimpleNamespace(Count=0)

        @property
        def CommentsThreaded(self):
            error = com_error(hresult, message)
            error.hresult = hresult
            raise error

    assert has_removable_artifacts(LegacySheet()) is False
    delete_removable_artifacts(LegacySheet())


def test_missing_threaded_comment_member_is_treated_as_absent() -> None:
    sheet = SimpleNamespace(
        Comments=SimpleNamespace(Count=0), Shapes=SimpleNamespace(Count=0)
    )

    assert has_removable_artifacts(sheet) is False
    delete_removable_artifacts(sheet)


def test_unsupported_threaded_comment_count_is_treated_as_absent_during_deletion() -> None:
    com_error = type("com_error", (Exception,), {})

    class UnsupportedThreadedComments:
        @property
        def Count(self) -> int:
            error = com_error(-2147352573, "member not found")
            error.hresult = -2147352573
            raise error

    sheet = SimpleNamespace(
        Comments=SimpleNamespace(Count=0),
        CommentsThreaded=UnsupportedThreadedComments(),
        Shapes=SimpleNamespace(Count=0),
    )

    delete_removable_artifacts(sheet)


def test_unexpected_threaded_comment_error_propagates() -> None:
    marker = RuntimeError("COM transport lost")

    class BrokenSheet:
        Comments = SimpleNamespace(Count=0)
        Shapes = SimpleNamespace(Count=0)

        @property
        def CommentsThreaded(self):
            raise marker

    with pytest.raises(RuntimeError) as caught:
        has_removable_artifacts(BrokenSheet())
    assert caught.value is marker


def test_deletion_walks_each_collection_backwards_until_empty() -> None:
    comments = MutableCollection(("legacy-1", "legacy-2", "legacy-3"))
    threaded = MutableCollection(("threaded-1", "threaded-2"))
    shapes = MutableCollection(("shape-1", "shape-2", "shape-3"))
    sheet = SimpleNamespace(
        Comments=comments, CommentsThreaded=threaded, Shapes=shapes
    )

    delete_removable_artifacts(sheet)

    assert comments.deleted == ["legacy-3", "legacy-2", "legacy-1"]
    assert threaded.deleted == ["threaded-2", "threaded-1"]
    assert shapes.deleted == ["shape-3", "shape-2", "shape-1"]
    assert (comments.Count, threaded.Count, shapes.Count) == (0, 0, 0)
