param(
    [string]$PythonExe = "",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $ProjectRoot
if (-not $PythonExe) {
    $PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
}

$ProjectMarker = Join-Path $ProjectRoot "pyproject.toml"
$EntryPoint = Join-Path $ProjectRoot "src\excel_splitter\__main__.py"
$AppIcon = Join-Path $ProjectRoot "src\excel_splitter\assets\app.ico"
if (-not (Test-Path -LiteralPath $ProjectMarker -PathType Leaf) -or
    -not (Test-Path -LiteralPath $EntryPoint -PathType Leaf) -or
    -not (Test-Path -LiteralPath $AppIcon -PathType Leaf)) {
    throw "Could not validate the Excel File Toolkit project root: $ProjectRoot"
}

function Invoke-Checked {
    param([scriptblock]$Command, [string]$Description)
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE"
    }
}

if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf)) {
    throw "Python executable not found: $PythonExe"
}

Invoke-Checked { & $PythonExe -m pip check } "Dependency verification"
Invoke-Checked {
    & $PythonExe -c "from importlib.metadata import version; expected={'pywin32':'312','pytest':'8.4.2','pyinstaller':'6.22.2'}; actual={k:version(k) for k in expected}; assert actual == expected, f'Pinned dependency mismatch: {actual}'"
} "Pinned-version verification"
if (-not $SkipTests) {
    Invoke-Checked { & $PythonExe -m pytest -q } "Regression tests"
}

$PythonComDll = (& $PythonExe -c "import pythoncom; print(pythoncom.__file__)").Trim()
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $PythonComDll -PathType Leaf)) {
    throw "pythoncom native DLL was not found: $PythonComDll"
}
if ([System.IO.Path]::GetFileName($PythonComDll) -notmatch '^pythoncom\d+\.dll$') {
    throw "Unexpected pythoncom native DLL name: $PythonComDll"
}
$PythonComBinary = "$PythonComDll;pywin32_system32"

Invoke-Checked {
    & $PythonExe -m PyInstaller --noconfirm --clean --onefile --windowed `
        --name ExcelFileToolkit --paths src --add-binary $PythonComBinary `
        --icon $AppIcon --add-data "$AppIcon;excel_splitter/assets" `
        src/excel_splitter/__main__.py
} "One-file build"

$FinalExe = Join-Path $ProjectRoot "dist\ExcelFileToolkit.exe"
if (-not (Test-Path -LiteralPath $FinalExe -PathType Leaf)) {
    throw "Final executable was not created: $FinalExe"
}
$ArchiveListing = & $PythonExe -m PyInstaller.utils.cliutils.archive_viewer `
    -r -b $FinalExe
if ($LASTEXITCODE -ne 0) {
    throw "Could not inspect the final executable archive."
}
if (-not ($ArchiveListing -match 'pywin32_system32[\\/]pythoncom\d+\.dll')) {
    throw "Final executable omitted the pythoncom native DLL."
}
if (-not ($ArchiveListing -match 'excel_splitter[\\/]assets[\\/]app\.ico')) {
    throw "Final executable omitted the window icon."
}
$SelfTest = Start-Process -FilePath $FinalExe `
    -ArgumentList "--self-test-pywin32" -WindowStyle Hidden -PassThru
if (-not $SelfTest.WaitForExit(30000)) {
    Stop-Process -Id $SelfTest.Id -Force -ErrorAction SilentlyContinue
    throw "Final executable pywin32 self-test timed out after 30 seconds."
}
if ($SelfTest.ExitCode -ne 0) {
    throw "Final executable could not import pythoncom/win32com (exit $($SelfTest.ExitCode))."
}
Invoke-Checked {
    & $PythonExe scripts/check_executable_icon.py $FinalExe
} "Executable and window icon verification"
Write-Host "Built: $FinalExe"
