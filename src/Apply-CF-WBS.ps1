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

$excel = $null
$wb = $null
$ws = $null

$m = [System.Reflection.Missing]::Value

try {
  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  } catch {
    $excel = New-Object -ComObject Excel.Application 
  }
  $prevStyle = $excel.ReferenceStyle
  $excel.ReferenceStyle = -4150 # xlR1C1
  $excel.ScreenUpdating = $false

  if ($WorkbookPath) {
    $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path) 
  } else {
    $wb = $excel.ActiveWorkbook 
  }
  $ws = if ($SheetName) {
    $wb.Worksheets.Item($SheetName) 
  } else {
    $excel.ActiveSheet 
  }
  $sep = $excel.International(5)

  # --- Cleanup ---
  $rChart = $ws.Range($ws.Cells($HeaderRow, $StartCol), $ws.Cells($EndRow, $EndCol))
  $rD     = $ws.Range($ws.Cells($StartRowBands, 5), $ws.Cells($EndRow, 5))
  $rR     = $ws.Range($ws.Cells($StartRowBands, 2), $ws.Cells($EndRow, 8))
  foreach ($r in @($rChart, $rD, $rR)) {
    try {
      $r.FormatConditions.Delete() 
    } catch {
    } 
  }

  # Formula Parts
  $idx = "R" + $HeaderRow + "C"
  $rowA = "RC1"; $rowE = "RC5"; $rowF = "RC6"; $rowG = "RC7"
  $isBand = "(" + $rowG + ">=" + $idx + ")*(" + $rowF + "<" + $idx + "+1)"

  # --- Add Rules (Order: Lowest Priority to Highest Priority) ---

  # 1. Default Band (Blue)
  $fcBand = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $isBand + ")")
  $fcBand.Interior.ThemeColor = 8 # Accent4
  $fcBand.StopIfTrue = $false

  # 2. Importance Bands (3, 2, 1)
  # A=3 (Light Yellow)
  $fc3 = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $rowA + "=3" + $sep + $isBand + ")")
  $fc3.Interior.Color = 13172735 # RGB(255, 255, 200)
  $fc3.StopIfTrue = $true

  # A=2 (Light Orange)
  $fc2 = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $rowA + "=2" + $sep + $isBand + ")")
  $fc2.Interior.Color = 11853055 # RGB(255, 220, 180)
  $fc2.StopIfTrue = $true

  # A=1 (Vivid Red/Pink)
  $fc1 = $rChart.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $rowA + "=1" + $sep + $isBand + ")")
  $fc1.Interior.Color = 7895295 # RGB(255, 120, 120) - Darker Red/Pink
  $fc1.StopIfTrue = $true

  # 3. Progress (Gray)
  $fProg = "=AND(ROW()>=" + $StartRowBands + $sep + $rowF + "<=" + $idx + $sep + "ROUNDDOWN((" + $rowG + "-" + $rowF + "+1)*" + $rowE + $sep + "0)+" + $rowF + "-1>=" + $idx + ")"
  $fcProg = $rChart.FormatConditions.Add(2, $m, $fProg)
  $fcProg.Interior.ThemeColor = 1; $fcProg.Interior.TintAndShade = -0.35
  $fcProg.StopIfTrue = $false

  # 4. Baseline Line (Red border, very light gray bg)
  $fcBase = $rChart.FormatConditions.Add(2, $m, "=AND(" + $BaselineAddress + ">=" + $idx + $sep + $BaselineAddress + "<" + $idx + "+1)")
  $fcBase.Interior.Color = 15790320 # RGB(240, 240, 240)
  try {
    foreach ($i in @(7, 10)) {
      $b = $fcBase.Borders.Item($i)
      $b.LineStyle = 1; $b.Weight = 2; $b.Color = 255 # Bold Red
    }
  } catch {
  }
  $fcBase.StopIfTrue = $false

  # 5. Row Rule (Highlights delayed tasks in Red)
  $fRow = "=AND(" + $rowE + "<1" + $sep + $rowG + "<" + $BaselineAddress + $sep + $rowG + "<>"""")"
  $fcRow = $rR.FormatConditions.Add(2, $m, $fRow)
  $fcRow.Interior.Color = 13551615; $fcRow.Font.Bold = $true; $fcRow.Font.Color = 393372

  # 6. Gray out completed tasks (Progress = 100%)
  $fcDone = $rR.FormatConditions.Add(2, $m, "=" + $rowE + "=1")
  $fcDone.Interior.Color = 14211288; $fcDone.Font.Color = 10526880

  # --- Force Priority ---
  $rules = @($fcBand, $fc3, $fc2, $fc1, $fcProg, $fcBase)
  foreach ($rule in $rules) {
    if ($null -ne $rule) {
      try {
        $rule.SetFirstPriority() 
      } catch {
      } 
    } 
  }

  # --- DataBar ---
  $fcBar = $rD.FormatConditions.AddDatabar()
  try {
    $fcBar.MinPoint.Modify(0, 0); $fcBar.MaxPoint.Modify(0, 1) 
  } catch {
  }

  # --- Grouping ---
  $vals = $ws.Range($ws.Cells($StartRowBands, 5), $ws.Cells($EndRow, 5)).Value2
  for ($i = 1; $i -le $vals.Length; $i++) {
    $rowIdx = $StartRowBands + $i - 1; $prog = $vals[$i, 1]
    $ind = 0; try {
      $ind = $ws.Cells($rowIdx, 2).IndentLevel 
    } catch {
    }
    if ($null -ne $prog -and $prog -eq 1) {
      $ws.Rows($rowIdx).OutlineLevel = 4 
    } else {
      $ws.Rows($rowIdx).OutlineLevel = [math]::Min(3, $ind + 1) 
    }
  }

  $excel.ReferenceStyle = $prevStyle
  $excel.ScreenUpdating = $true
  Write-Host "Success."
} catch {
  if ($null -ne $excel) {
    try {
      $excel.ReferenceStyle = $prevStyle; $excel.ScreenUpdating = $true 
    } catch {
    } 
  }
  Write-Error $_.Exception.Message
} finally {
  foreach ($o in @($ws, $wb, $excel)) {
    if ($null -ne $o) {
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null 
    } 
  }
}
