"""Check packaged ICO resources and the launched EXE's actual window icons."""

from pathlib import Path
import struct
import subprocess
import sys
import time

import win32api
import win32con
import win32gui
import win32job
import win32process
import win32ui


ICON = Path(__file__).resolve().parents[1] / "src/excel_splitter/assets/app.ico"


def check_resources(executable: Path) -> None:
    data = ICON.read_bytes()
    count = struct.unpack_from("<H", data, 4)[0]
    module = win32api.LoadLibraryEx(str(executable), 0, win32con.LOAD_LIBRARY_AS_DATAFILE)
    try:
        groups = win32api.EnumResourceNames(module, win32con.RT_GROUP_ICON)
        group = win32api.LoadResource(module, win32con.RT_GROUP_ICON, groups[0])
        assert struct.unpack_from("<HHH", group) == (0, 1, count), "Icon sizes missing"
        for index in range(count):
            entry = data[6 + 16 * index:22 + 16 * index]
            size, offset = struct.unpack_from("<II", entry, 8)
            resource_id = struct.unpack_from("<H", group, 18 + 14 * index)[0]
            assert group[6 + 14 * index:18 + 14 * index] == entry[:12]
            assert win32api.LoadResource(module, win32con.RT_ICON, resource_id) == data[offset:offset + size]
    finally:
        win32api.FreeLibrary(module)


def _pixels(handle):
    info = win32gui.GetIconInfo(handle)
    bitmap = win32ui.CreateBitmapFromHandle(info[4])
    dimensions = bitmap.GetInfo()
    return dimensions["bmWidth"], dimensions["bmHeight"], bitmap.GetBitmapBits(True)


def check_window(executable: Path) -> None:
    # Own the test process tree so cleanup cannot close the user's other apps.
    job = win32job.CreateJobObject(None, "")
    limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    limits["BasicLimitInformation"]["LimitFlags"] = win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
    startup = subprocess.STARTUPINFO()
    startup.dwFlags = subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = win32con.SW_HIDE
    process = subprocess.Popen([str(executable)], startupinfo=startup)
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, process.pid)
        try:
            win32job.AssignProcessToJobObject(job, handle)
        finally:
            handle.Close()
        deadline = time.monotonic() + 30
        windows = []
        while time.monotonic() < deadline:
            pids = win32job.QueryInformationJobObject(job, win32job.JobObjectBasicProcessIdList)
            win32gui.EnumWindows(
                lambda hwnd, _: windows.append(hwnd)
                if win32process.GetWindowThreadProcessId(hwnd)[1] in pids
                and win32gui.GetWindowText(hwnd) == "Excel File Toolkit" else None, None)
            if windows:
                break
            assert process.poll() is None, "EXE exited before creating its window"
            time.sleep(0.1)
        assert len(windows) == 1, "Toolkit window not found"
        window = windows[0]
        for kind in (win32con.ICON_SMALL, win32con.ICON_BIG):
            actual = win32gui.SendMessage(window, win32con.WM_GETICON, kind, 0)
            assert actual, "Window icon missing"
            width, height, pixels = _pixels(actual)
            expected = win32gui.LoadImage(None, str(ICON), win32con.IMAGE_ICON, width, height,
                                         win32con.LR_LOADFROMFILE)
            try:
                assert pixels == _pixels(expected)[2], "Window icon differs from selected artwork"
            finally:
                win32gui.DestroyIcon(expected)
        win32gui.PostMessage(window, win32con.WM_CLOSE, 0, 0)
        assert process.wait(timeout=15) == 0, "EXE did not close normally"
    finally:
        job.Close()
        if process.poll() is None:
            process.kill()
        process.wait(timeout=5)


if __name__ == "__main__":
    for argument in sys.argv[1:] or ["dist/ExcelFileToolkit.exe"]:
        executable = Path(argument).resolve(strict=True)
        check_resources(executable)
        check_window(executable)
        print(f"PASS: embedded ICO and live small/large window icons: {executable}", flush=True)
