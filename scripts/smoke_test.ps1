param(
    [Parameter(Mandatory = $true)][string]$ExePath,
    [Parameter(Mandatory = $true)][string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$ExePath = [System.IO.Path]::GetFullPath($ExePath)
$WorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)

if (-not (Test-Path -LiteralPath $ExePath -PathType Leaf)) {
    throw "Executable not found: $ExePath"
}
if (Test-Path -LiteralPath $WorkbookPath) {
    throw "Fixture path already exists; refusing to overwrite it: $WorkbookPath"
}

$Excel = $null
$Workbook = $null
try {
    try {
        $Excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Desktop Microsoft Excel is required for this smoke test."
    }
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Add()
    $Sheet = $Workbook.Worksheets.Item(1)
    $Sheet.Name = "분류표"
    while ($Workbook.Worksheets.Count -gt 1) {
        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
    }
    $Reference = $Workbook.Worksheets.Add()
    $Reference.Name = "참조"
    $Reference.Move($null, $Sheet)
    $Sheet.Activate()

    $Sheet.Range("A1").Value2 = "구분"
    $Sheet.Range("B1").Value2 = "금액"
    $Sheet.Range("C1").Value2 = "계산"
    $Sheet.Range("A2").Value2 = "A"
    $Sheet.Range("B2").Value2 = 10
    $Sheet.Range("A3").Value2 = "A"
    $Sheet.Range("B3").Value2 = 20
    $Sheet.Range("B4").Value2 = 30
    $Table = $Sheet.ListObjects.Add(1, $Sheet.Range("A1:C4"), $null, 1)
    $Table.Name = "Table1"
    $Table.ListColumns.Item("계산").DataBodyRange.Formula = "=[@금액]*2"
    $Table.ShowTotals = $true
    $Table.ListColumns.Item("금액").TotalsCalculation = 1
    $Sheet.Range("E2").Formula = "=SUM(Table1[금액])"
    $Sheet.Range("E5:F5").Merge()
    $Sheet.Range("E5").Value2 = "Table 밖 병합"
    $Sheet.Columns.Item("A").ColumnWidth = 18
    $Sheet.Rows.Item(2).RowHeight = 24
    $Sheet.Range("B2:B4").FormatConditions.AddColorScale(2) | Out-Null
    $Sheet.Range("A2:A4").Validation.Add(3, 1, 1, "A,B")
    $Sheet.Shapes.AddShape(1, 320, 20, 90, 35).TextFrame.Characters().Text = "fixture"
    $ChartObject = $Sheet.ChartObjects().Add(320, 80, 260, 160)
    $ChartObject.Chart.SetSourceData($Sheet.Range("B1:B4"))
    $Sheet.PageSetup.Orientation = 2
    $Reference.Range("A1").Value2 = "참조 시트"

    $FixtureDir = Split-Path -Parent $WorkbookPath
    [System.IO.Directory]::CreateDirectory($FixtureDir) | Out-Null
    $Workbook.SaveAs($WorkbookPath, 51)
    $Workbook.Close($false)
    $Workbook = $null
}
finally {
    if ($null -ne $Workbook) { $Workbook.Close($false) }
    if ($null -ne $Excel) { $Excel.Quit() }
    if ($null -ne $Workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) }
    if ($null -ne $Excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$SourceHash = (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
Write-Host "Fixture oracle:"
Write-Host "  visible sheets: 분류표, 참조"
Write-Host "  table: 분류표!Table1; classification column: 구분"
Write-Host "  rows: A, A, blank"
Write-Host "  pattern: %_결과"
Write-Host "  A_결과.xlsx: 2 data rows; only 분류표"
Write-Host "  _결과.xlsx: 1 data row; only 분류표"
Write-Host "  source SHA-256 before: $SourceHash"
Write-Host "Manual checklist: calculated column, formula, total row, conditional format, validation, merge, width/height, shape, chart, and print setting are preserved."
Write-Host "Confirm the deleted-sheet-reference warning appears and no orphan Excel process remains."

$Process = Start-Process -FilePath $ExePath -PassThru
Read-Host "Complete the split in the app, close it, then press Enter"
if (-not $Process.HasExited) {
    throw "Close ExcelSplitter before completing the smoke test."
}
$AfterHash = (Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
if ($AfterHash -ne $SourceHash) {
    throw "Source workbook hash changed. Before=$SourceHash After=$AfterHash"
}
Write-Host "Source SHA-256 unchanged: $AfterHash"
