param(
    [string]$PythonExe = "C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $ProjectRoot

$ProjectMarker = Join-Path $ProjectRoot "pyproject.toml"
$EntryPoint = Join-Path $ProjectRoot "src\excel_splitter\__main__.py"
if (-not (Test-Path -LiteralPath $ProjectMarker -PathType Leaf) -or
    -not (Test-Path -LiteralPath $EntryPoint -PathType Leaf)) {
    throw "Could not validate the ExcelSplitter project root: $ProjectRoot"
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
Invoke-Checked { & $PythonExe -m pytest tests/unit -q } "Unit tests"

$PythonComDll = (& $PythonExe -c "import pythoncom; print(pythoncom.__file__)").Trim()
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $PythonComDll -PathType Leaf)) {
    throw "pythoncom native DLL was not found: $PythonComDll"
}
if ([System.IO.Path]::GetFileName($PythonComDll) -notmatch '^pythoncom\d+\.dll$') {
    throw "Unexpected pythoncom native DLL name: $PythonComDll"
}
$PythonComBinary = "$PythonComDll;pywin32_system32"

Invoke-Checked {
    & $PythonExe -m PyInstaller --noconfirm --clean --onedir --windowed `
        --name ExcelSplitter --paths src --add-binary $PythonComBinary `
        src/excel_splitter/__main__.py
} "One-folder build"

$OnedirPath = (Join-Path $ProjectRoot "dist\ExcelSplitter")
$OnedirExe = Join-Path $OnedirPath "ExcelSplitter.exe"
if (-not (Test-Path -LiteralPath $OnedirExe -PathType Leaf)) {
    throw "One-folder executable was not created: $OnedirExe"
}
$OnedirPythonCom = Get-ChildItem -LiteralPath $OnedirPath -Recurse -File `
    -Filter "pythoncom*.dll"
if (-not $OnedirPythonCom) {
    throw "One-folder build omitted the pythoncom native DLL."
}
$Probe = Start-Process -FilePath $OnedirExe -PassThru
Start-Sleep -Seconds 2
if ($Probe.HasExited) {
    throw "One-folder executable exited during startup verification."
}
$Probe.CloseMainWindow() | Out-Null
if (-not $Probe.WaitForExit(3000)) {
    Stop-Process -Id $Probe.Id
}

$BuildPath = Join-Path $ProjectRoot "build\ExcelSplitter"
Remove-Item -LiteralPath $OnedirPath -Recurse -Force
if (Test-Path -LiteralPath $BuildPath) {
    Remove-Item -LiteralPath $BuildPath -Recurse -Force
}

Invoke-Checked {
    & $PythonExe -m PyInstaller --noconfirm --clean --onefile --windowed `
        --name ExcelSplitter --paths src --add-binary $PythonComBinary `
        src/excel_splitter/__main__.py
} "One-file build"

$FinalExe = Join-Path $ProjectRoot "dist\ExcelSplitter.exe"
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
$SelfTest = Start-Process -FilePath $FinalExe `
    -ArgumentList "--self-test-pywin32" -WindowStyle Hidden -PassThru
if (-not $SelfTest.WaitForExit(30000)) {
    Stop-Process -Id $SelfTest.Id -Force -ErrorAction SilentlyContinue
    throw "Final executable pywin32 self-test timed out after 30 seconds."
}
if ($SelfTest.ExitCode -ne 0) {
    throw "Final executable could not import pythoncom/win32com (exit $($SelfTest.ExitCode))."
}
Write-Host "Built: $FinalExe"
