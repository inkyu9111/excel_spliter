from __future__ import annotations

import threading
import logging
import tkinter as tk

from .controller import AppController
from .excel_gateway import ExcelComGateway
from .toolkit_gui import ExcelFileToolkitGui
from .ports import SplitServicePort
from .split_service import SplitService


def _prewarm_gateway(gateway: ExcelComGateway) -> None:
    try:
        gateway.prewarm()
    except Exception:
        # The first real action retries this startup path and shows the
        # existing localized error; bootstrap must not terminate the GUI.
        logging.getLogger(__name__).warning("Excel 사전 시작 실패; 다음 작업에서 재시도합니다.", exc_info=True)


def main(service: SplitServicePort | None = None) -> None:
    gateway: ExcelComGateway | None = None
    if service is None:
        gateway = ExcelComGateway()
        concrete: SplitServicePort = SplitService(gateway)
    else:
        concrete = service
    root = tk.Tk()
    ExcelFileToolkitGui(root, AppController(concrete))
    if gateway is not None:
        threading.Thread(
            target=_prewarm_gateway,
            args=(gateway,),
            name="excel-prewarm",
            daemon=True,
        ).start()
    root.mainloop()
