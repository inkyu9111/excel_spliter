import sys

from excel_splitter.app import main


def run() -> int:
    if sys.argv[1:] == ["--self-test-pywin32"]:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401

        return 0

    main()
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
