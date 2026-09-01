from __future__ import annotations

import tkinter as tk

from .controller import AppController
from .excel_gateway import ExcelComGateway
from .gui import ExcelSplitterGui
from .ports import SplitServicePort
from .split_service import SplitService


def main(service: SplitServicePort | None = None) -> None:
    concrete = service or SplitService(ExcelComGateway())
    root = tk.Tk()
    ExcelSplitterGui(root, AppController(concrete))
    root.mainloop()
