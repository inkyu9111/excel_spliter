"""Run with EXCEL_E2E=1; requires Windows, desktop Excel and pywin32.

Each case runs in a fresh Python process. Keep this suite serial: Excel's
clipboard is shared by the Split/Merge workflows.
"""

import os
from pathlib import Path
import subprocess
import sys

import pytest


ROOT = Path(__file__).resolve().parents[2]
pytestmark = pytest.mark.skipif(
    os.environ.get("EXCEL_E2E") != "1",
    reason="Set EXCEL_E2E=1 to run the installed desktop Excel workflows",
)


@pytest.mark.parametrize("script", (
    "check_merge_excel.py",
    "check_split_merge_excel.py",
    "check_compare_excel.py",
    "check_etc_excel.py",
    "check_gui_excel.py",
))
def test_native_excel_workflow(script, monkeypatch):
    assert sys.platform == "win32", "Desktop Excel E2E requires Windows"
    assert not os.environ.get("PYTEST_XDIST_WORKER"), "Run Excel E2E serially without pytest-xdist"
    monkeypatch.setenv("PYTHONUTF8", "1")
    subprocess.run(
        [sys.executable, "-u", str(ROOT / "scripts" / script)],
        cwd=ROOT,
        timeout=900,
        check=True,
    )
