from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_splitter import gui


@pytest.fixture
def app_logger():
    logger = logging.getLogger("excel_splitter")
    original_handlers = tuple(logger.handlers)
    original_level = logger.level
    original_propagate = logger.propagate
    for handler in original_handlers:
        logger.removeHandler(handler)
    logger.setLevel(logging.NOTSET)
    logger.propagate = False
    try:
        yield logger
    finally:
        for handler in tuple(logger.handlers):
            logger.removeHandler(handler)
            handler.close()
        for handler in original_handlers:
            logger.addHandler(handler)
        logger.setLevel(original_level)
        logger.propagate = original_propagate


def test_window_close_releases_its_log_and_can_reopen_it(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, app_logger: logging.Logger
) -> None:
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    logger = gui.configure_logging()
    log_path = gui._log_path(logger)
    foreign_handler = logging.StreamHandler()
    app_logger.addHandler(foreign_handler)
    destroyed: list[bool] = []

    def shutdown() -> None:
        raise RuntimeError("stop")

    window = object.__new__(gui.ExcelSplitterGui)
    window._busy = False
    window.controller = SimpleNamespace(shutdown=shutdown)
    window.root = SimpleNamespace(destroy=lambda: destroyed.append(True))
    window.logger = logger

    window._on_close()

    released_path = log_path.with_suffix(".released")
    log_path.rename(released_path)
    reopened_logger = gui.configure_logging()
    owned_handlers = [
        handler
        for handler in reopened_logger.handlers
        if isinstance(handler, RotatingFileHandler)
    ]
    reopened_logger.info("reopened")
    owned_handlers[0].flush()

    assert destroyed == [True]
    assert foreign_handler in reopened_logger.handlers
    assert len(owned_handlers) == 1
    assert "reopened" in log_path.read_text(encoding="utf-8")
