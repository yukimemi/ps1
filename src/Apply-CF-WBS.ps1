<#
.SYNOPSIS
  Excel WBS Conditional Formatting & Grouping Tool (R1C1 Style)
#>

[CmdletBinding()]
param(
  [string]$WorkbookPath,
  [switch]$UseActive,
  [string]$SheetName       = $null,
  [string]$BaselineAddress = "R3C6", # $F$3
  [int]$HeaderRow          = 5,
  [int]$StartCol           = 10,   # J
  [int]$EndCol             = 65,   # BM
  [int]$StartRowBands      = 11,
  [int]$EndRow             = 1000,
  [switch]$SaveChanges
)

$scriptStartTime = Get-Date
function Log-Progress {
  param([string]$message)
  $elapsed = (Get-Date) - $scriptStartTime
  Write-Host ("[{0:hh\:mm\:ss}] {1}" -f $elapsed, $message)
}

$excel = $null; $wb = $null; $ws = $null; $attached = $false
$m = [System.Reflection.Missing]::Value

try {
  Log-Progress "Connecting to Excel..."
  try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application"); $attached = $true } catch { $excel = New-Object -ComObject Excel.Application }

  Log-Progress "Setting reference style and disabling updates..."
  $prevStyle = $excel.ReferenceStyle
  try {
    $prevCalc = $excel.Calculation
    $excel.Calculation = -4135 # xlManual
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.ReferenceStyle = -4150 # xlR1C1
  } catch {}

  Log-Progress "Accessing Sheet..."
  if ($WorkbookPath) { $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path) } else { $wb = $excel.ActiveWorkbook }
  $ws = if ($SheetName) { $wb.Worksheets.Item($SheetName) } else { $excel.ActiveSheet }
  $sep = $excel.International(5)

  Log-Progress "Cleaning up existing formats..."
  $rChart = $ws.Range($ws.Cells($HeaderRow, $StartCol), $ws.Cells($EndRow, $EndCol))
  $rD     = $ws.Range($ws.Cells($StartRowBands, 5), $ws.Cells($EndRow, 5))
  $rR     = $ws.Range($ws.Cells($StartRowBands, 2), $ws.Cells($EndRow, 8))
  foreach ($r in @($rChart, $rD, $rR)) { try { $r.FormatConditions.Delete() } catch {} }

  Log-Progress "Adding rules (Using Selection for Baseline robustness)..."
  $idx = "R" + $HeaderRow + "C"
  $rowA = "RC1"; $rowE = "RC5"; $rowF = "RC6"; $rowG = "RC7"
  $isBand = "(" + $rowG + ">=" + $idx + ")*(" + $rowF + "<" + $idx + "+1)"

  # 1. Band rules
  $fcB = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $isBand + ")")
  $fcB.Interior.ThemeColor = 8
  foreach ($imp in @(@(3, 13172735), @(2, 11853055), @(1, 7895295))) {
    $fcI = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $rowA + "=" + $imp[0] + $sep + $isBand + ")")
    $fcI.Interior.Color = $imp[1]; $fcI.StopIfTrue = $true
  }

  # 2. Progress
  $fcP = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $rowF + "<=" + $idx + $sep + "ROUNDDOWN((" + $rowG + "-" + $rowF + "+1)*" + $rowE + $sep + "0)+" + $rowF + "-1>=" + $idx + ")")
  $fcP.Interior.ThemeColor = 1; $fcP.Interior.TintAndShade = -0.35

  # 3. Baseline (Selection Approach)
  $fBase = "=AND(" + $BaselineAddress + ">=" + $idx + $sep + $BaselineAddress + "<" + $idx + "+1)"
  $fcBase = $rChart.FormatConditions.Add(2, $m, $fBase)
  $fcBase.SetFirstPriority()
  
  # Index 1 is now the Baseline rule
  $fc1 = $rChart.FormatConditions.Item(1)
  try {
    $fc1.Borders.Item(-4131).LineStyle = 1
    $fc1.Borders.Item(-4131).Color = 255
    $fc1.Borders.Item(-4131).TintAndShade = 0
    $fc1.Borders.Item(-4131).Weight = 2
  } catch {
    # If Item(7) fails, try setting border on the whole object
    Log-Progress "Warning: Selection border failed. Trying fallback background color..."
    $fc1.Interior.Color = 13551615
  }
  $fc1.StopIfTrue = $false

  # 4. Row Rules
  $fcRow = $rR.FormatConditions.Add(2, $m, "=AND(" + $rowE + "<1" + $sep + $rowG + "<" + $BaselineAddress + $sep + $rowG + "<>"""")")
  $fcRow.Interior.Color = 13551615; $fcRow.Font.Bold = $true; $fcRow.Font.Color = 393372
  $fcDone = $rR.FormatConditions.Add(2, $m, "=" + $rowE + "=1")
  $fcDone.Interior.Color = 14211288; $fcDone.Font.Color = 10526880

  # Priority Adjustment
  try { $fcP.SetFirstPriority() } catch {}
  try { $fc1.SetFirstPriority() } catch {}

  # 5. DataBar
  $fcBar = $rD.FormatConditions.AddDatabar()
  try { $fcBar.MinPoint.Modify(0, 0); $fcBar.MaxPoint.Modify(0, 1) } catch {}

  Log-Progress "Grouping..."
  $progVals = $ws.Range($ws.Cells($StartRowBands, 5), $ws.Cells($EndRow, 5)).Value2
  for ($i = 1; $i -le $progVals.Length; $i++) {
    $rowIdx = $StartRowBands + $i - 1; $p = $progVals[$i, 1]
    $level = 1
    if ($null -ne $p -and $p -eq 1) { $level = 4 }
    else { try { $level = [math]::Min(3, $ws.Cells($rowIdx, 2).IndentLevel + 1) } catch { $level = 1 } }
    if ($ws.Rows($rowIdx).OutlineLevel -ne $level) { $ws.Rows($rowIdx).OutlineLevel = $level }
  }

  Log-Progress "Finalizing..."
  if ($null -ne $prevCalc) { try { $excel.Calculation = $prevCalc } catch {} }
  try { $excel.ReferenceStyle = $prevStyle; $excel.ScreenUpdating = $true; $excel.EnableEvents = $true } catch {}
  if ($SaveChanges) { $wb.Save() }
  Log-Progress "Success."

} catch {
  Log-Progress "Error: $($_.Exception.Message)"
  if ($null -ne $excel) { try { $excel.Calculation = $prevCalc; $excel.ReferenceStyle = $prevStyle; $excel.ScreenUpdating = $true; $excel.EnableEvents = $true } catch {} }
  Write-Error $_.Exception.Message
} finally {
  foreach ($o in @($ws, $wb, $excel)) { if ($null -ne $o) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } }
}
