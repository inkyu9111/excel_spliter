from __future__ import annotations

from typing import Any


def unsupported_threaded_comments_error(exc: Exception) -> bool:
    if isinstance(exc, AttributeError):
        return True
    hresult = getattr(exc, "hresult", exc.args[0] if exc.args else None)
    if exc.__class__.__name__ != "com_error":
        return False
    if hresult in (-2147352573, -2147352570):
        return True
    message = str(exc).casefold()
    return hresult == -2147352567 and (
        "commentthreaded" in message
        or "threaded comment" in message
        or "스레드 주석" in message
    )


def legacy_comment_count(sheet: Any) -> int:
    return int(sheet.Comments.Count)


def _threaded_comments(sheet: Any) -> Any | None:
    try:
        return getattr(sheet, "CommentsThreaded")
    except Exception as exc:
        if unsupported_threaded_comments_error(exc):
            return None
        raise


def threaded_comment_count(sheet: Any) -> int:
    collection = _threaded_comments(sheet)
    if collection is None:
        return 0
    try:
        return int(collection.Count)
    except Exception as exc:
        if unsupported_threaded_comments_error(exc):
            return 0
        raise


def has_removable_artifacts(sheet: Any) -> bool:
    return (
        legacy_comment_count(sheet) > 0
        or threaded_comment_count(sheet) > 0
        or int(sheet.Shapes.Count) > 0
    )


def delete_collection_backwards(collection: Any) -> None:
    for index in range(int(collection.Count), 0, -1):
        collection.Item(index).Delete()


def delete_supported_threaded_comments_backwards(sheet: Any) -> None:
    collection = _threaded_comments(sheet)
    if collection is None:
        return
    try:
        count = int(collection.Count)
    except Exception as exc:
        if unsupported_threaded_comments_error(exc):
            return
        raise
    for index in range(count, 0, -1):
        collection.Item(index).Delete()


def delete_removable_artifacts(sheet: Any) -> None:
    delete_collection_backwards(sheet.Comments)
    delete_supported_threaded_comments_backwards(sheet)
    delete_collection_backwards(sheet.Shapes)
