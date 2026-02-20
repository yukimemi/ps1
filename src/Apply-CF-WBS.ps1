<#
.SYNOPSIS
  Excel WBS Conditional Formatting & Grouping Tool (R1C1 Style)
#>

[CmdletBinding()]
param(
  [string]$WorkbookPath,
  [switch]$UseActive,
  [string]$SheetName       = $null,
  [string]$BaselineAddress = "R3C6",
  [int]$HeaderRow          = 5,
  [int]$StartCol           = 10,
  [int]$EndCol             = 79,
  [int]$StartRowBands      = 11,
  [int]$EndRow             = 1000,
  [switch]$SaveChanges
)

$scriptStartTime = Get-Date
function Log-Progress {
  param([string]$m) $e = (Get-Date) - $scriptStartTime
  Write-Host ("[{0:hh\:mm\:ss}] {1}" -f $e, $m)
}

$excel = $null
$wb = $null
$ws = $null
$attached = $false
$m = [System.Reflection.Missing]::Value

try {
  Log-Progress "Connecting..."
  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $attached = $true
  } catch {
    $excel = New-Object -ComObject Excel.Application
  }
  $prevStyle = $excel.ReferenceStyle
  $excel.ReferenceStyle = -4150
  $excel.ScreenUpdating = $false
  $excel.Calculation = -4135

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

  # --- Define Holiday Name (Safe sheet naming) ---
  Log-Progress "Defining Holiday Range Name..."
  try {
    # Attempt to find sheet by name, handle potential encoding issues
    $wsHol = $null
    foreach ($sh in $wb.Worksheets) {
      if ($sh.Name -match "祝日") {
        $wsHol = $sh
        break
      }
    }
    if ($null -ne $wsHol) {
      $wb.Names.Add("HolidayList", $wsHol.Range("B:B")) | Out-Null
    } else {
      Log-Progress "Warning: '祝日' sheet not found."
    }
  } catch {
    Log-Progress "Error defining name: $($_.Exception.Message)"
  }

  Log-Progress "Cleanup..."
  $rChartFull = $ws.Range($ws.Cells($HeaderRow, $StartCol), $ws.Cells($EndRow, $EndCol))
  $rChartData = $ws.Range($ws.Cells(7, $StartCol), $ws.Cells($EndRow, $EndCol)) # Applying from Row 7
  $rD         = $ws.Range($ws.Cells($StartRowBands, 5), $ws.Cells($EndRow, 5))
  $rR         = $ws.Range($ws.Cells($StartRowBands, 2), $ws.Cells($EndRow, 8))

  [void]$ws.Cells.FormatConditions.Delete()

  Log-Progress "Adding rules..."
  $idx = "R" + $HeaderRow + "C"
  $rowA = "RC1"
  $rowE = "RC5"
  $rowF = "RC6"
  $rowG = "RC7"
  $isBand = "(" + $rowG + ">=" + $idx + ")*(" + $rowF + "<" + $idx + "+1)"
  $q = [char]34

  # 3. Progress
  $fcP = $rChartFull.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $rowF + "<=" + $idx + $sep + "ROUNDDOWN((" + $rowG + "-" + $rowF + "+1)*" + $rowE + $sep + "0)+" + $rowF + "-1>=" + $idx + ")")
  $fcP.Interior.ThemeColor = 1
  $fcP.Interior.TintAndShade = -0.35

  # 2. Bands (Apply to full chart, but with row filter)
  $imps = @(
    @(1, 7895295),
    @(2, 15624315),
    @(3, 3145645)
  )
  foreach ($imp in $imps) {
    $fImp = "=AND(ROW()>=" + $StartRowBands + $sep + $rowA + "=" + $imp[0] + $sep + $isBand + ")"
    $fcI = $rChartFull.FormatConditions.Add(2, $m, $fImp)
    $fcI.Interior.Color = $imp[1]
    $fcI.StopIfTrue = $false
  }
  $fcBand = $rChartFull.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $isBand + ")")
  $fcBand.Interior.ThemeColor = 8
  $fcBand.StopIfTrue = $false

  # expired
  $fcExpired = $rChartFull.FormatConditions.Add(2, $m, "=AND(ROW()>=" + $StartRowBands + $sep + $isBand + $sep + $rowE + "<1" + $sep + $rowG + "<" + $BaselineAddress + $sep + $rowG + "<>" + $q + $q + ")")
  $fcExpired.Interior.Color = 255
  $fcExpired.StopIfTrue = $false

  # 1. Holiday (Apply only to rChartData - from Row 7)
  $fHol = "=OR(WEEKDAY(" + $idx + $sep + "2)>=6" + $sep + "COUNTIF(HolidayList" + $sep + $idx + ")>0)"
  $fcHol = $rChartData.FormatConditions.Add(2, $m, $fHol)
  $fcHol.Interior.Color = 16118015
  $fcHol.StopIfTrue = $false

  # 4. Baseline
  $fcBase = $rChartFull.FormatConditions.Add(2, $m, "=AND(" + $BaselineAddress + ">=" + $idx + $sep + $BaselineAddress + "<" + $idx + "+1)")
  try {
    $bL = $fcBase.Borders.Item(-4131)
    $bL.LineStyle = 1
    $bL.Color = 255
    $bL.Weight = 2
  } catch {
  }
  $fcBase.StopIfTrue = $false

  # 5. Row Rules
  $fR = "=AND(" + $rowE + "<1" + $sep + $rowG + "<" + $BaselineAddress + $sep + $rowG + "<>" + $q + $q + ")"
  $fcRow = $rR.FormatConditions.Add(2, $m, $fR)
  $fcRow.Interior.Color = 13551615
  $fcRow.Font.Bold = $true
  $fcRow.Font.Color = 393372
  $fcRow.StopIfTrue = $false
  $fcDone = $rR.FormatConditions.Add(2, $m, "=" + $rowE + "=1")
  $fcDone.Interior.Color = 14211288
  $fcDone.Font.Color = 10526880

  Log-Progress "Priorities..."
  try {
    [void]$fcExpired.SetFirstPriority()
    [void]$fcRow.SetFirstPriority()
    [void]$fcDone.SetFirstPriority()
  } catch {
    Log-Progress "$($_ | Out-String)"
  }

  # 6. DataBar
  $fcBar = $rD.FormatConditions.AddDatabar()
  try {
    $fcBar.MinPoint.Modify(0, 0)
    $fcBar.MaxPoint.Modify(0, 1)
  } catch {
  }

  Log-Progress "Grouping..."
  $progVals = $ws.Range($ws.Cells($StartRowBands, 5), $ws.Cells($EndRow, 5)).Value2
  for ($i = 1; $i -le $progVals.Length; $i++) {
    $rowIdx = $StartRowBands + $i - 1
    $p = $progVals[$i, 1]
    $level = 1
    if ($null -ne $p -and $p -eq 1) {
      $level = 4
    } else {
      try {
        $level = [math]::Min(3, $ws.Cells($rowIdx, 2).IndentLevel + 1)
      } catch {
        $level = 1
      }
    }
    if ($ws.Rows($rowIdx).OutlineLevel -ne $level) {
      $ws.Rows($rowIdx).OutlineLevel = $level
    }
  }

  Log-Progress "Finalizing..."
  $excel.Calculation = -4105
  $excel.ReferenceStyle = $prevStyle
  $excel.ScreenUpdating = $true
  $excel.EnableEvents = $true
  if ($SaveChanges) {
    $wb.Save()
  }
  Log-Progress "Success."
} catch {
  Log-Progress "Error: $($_.Exception.Message)"
  if ($null -ne $excel) {
    try {
      $excel.Calculation = -4105
      $excel.ReferenceStyle = $prevStyle
      $excel.ScreenUpdating = $true
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
