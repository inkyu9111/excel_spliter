from __future__ import annotations

import tkinter as tk

from .controller import AppController
from .gui import ExcelSplitterGui
from .ports import SplitServicePort


def main(service: SplitServicePort | None = None) -> None:
    if service is None:
        raise RuntimeError("SplitService wiring is pending")
    root = tk.Tk()
    ExcelSplitterGui(root, AppController(service))
    root.mainloop()
