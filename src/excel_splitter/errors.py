class ExcelSplitterError(Exception):
    """Base error safe to translate at the GUI boundary."""


class ExcelUnavailableError(ExcelSplitterError):
    pass


class WorkbookValidationError(ExcelSplitterError):
    pass


class SplitExecutionError(ExcelSplitterError):
    pass
