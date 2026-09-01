param(
    [string]$PythonExe = "C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $ProjectRoot

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

Invoke-Checked {
    & $PythonExe -m PyInstaller --noconfirm --clean --onedir --windowed `
        --name ExcelSplitter --paths src src/excel_splitter/__main__.py
} "One-folder build"

$OnedirPath = (Join-Path $ProjectRoot "dist\ExcelSplitter")
$OnedirExe = Join-Path $OnedirPath "ExcelSplitter.exe"
if (-not (Test-Path -LiteralPath $OnedirExe -PathType Leaf)) {
    throw "One-folder executable was not created: $OnedirExe"
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
        --name ExcelSplitter --paths src src/excel_splitter/__main__.py
} "One-file build"

$FinalExe = Join-Path $ProjectRoot "dist\ExcelSplitter.exe"
if (-not (Test-Path -LiteralPath $FinalExe -PathType Leaf)) {
    throw "Final executable was not created: $FinalExe"
}
Write-Host "Built: $FinalExe"
