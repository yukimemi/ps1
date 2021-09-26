<#
  .SYNOPSIS
    excelmerge
  .DESCRIPTION
    excel を merge する。
  .INPUTS
    - path: config path
  .OUTPUTS
    - 0: SUCCESS / 1: ERROR
  .Last Change : 2021/08/21 18:54:23.
#>
param(
  [Parameter()]
  [string]$path
)

$ErrorActionPreference = "Stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire
# Enable-RunspaceDebug -BreakAll

<#
  .SYNOPSIS
    log
  .DESCRIPTION
    log message
  .INPUTS
    - msg
    - color
  .OUTPUTS
    - None
#>
function log {

  [CmdletBinding()]
  [OutputType([void])]
  param([string]$msg, [string]$color)
  trap { Write-Host "[log] Error $_"; throw $_ }

  $now = Get-Date -f "yyyy/MM/dd HH:mm:ss.fff"
  if ($color) {
    Write-Host -ForegroundColor $color "${now} ${msg}"
  }
  else {
    Write-Host "${now} ${msg}"
  }
}

<#
  .SYNOPSIS
    Init
  .DESCRIPTION
    Init
  .INPUTS
    - None
  .OUTPUTS
    - None
#>
function Start-Init {

  [CmdletBinding()]
  [OutputType([void])]
  param()
  trap { log "[Start-Init] Error $_" "Red"; throw $_ }

  log "[Start-Init] Start" "Cyan"

  $script:app = @{}

  $cmdFullPath = & {
    if ($env:__SCRIPTPATH) {
      return [System.IO.Path]::GetFullPath($env:__SCRIPTPATH)
    }
    else {
      return [System.IO.Path]::GetFullPath($script:MyInvocation.MyCommand.Path)
    }
  }
  $app.Add("cmdFile", $cmdFullPath)
  $app.Add("cmdDir", [System.IO.Path]::GetDirectoryName($app.cmdFile))
  $app.Add("cmdName", [System.IO.Path]::GetFileNameWithoutExtension($app.cmdFile))
  $app.Add("cmdFileName", [System.IO.Path]::GetFileName($app.cmdFile))

  $app.Add("pwd", [System.IO.Path]::GetFullPath((Get-Location).Path))

  $app.Add("now", (Get-Date -Format "yyyyMMddTHHmmssfffffff"))
  $logDir = [System.IO.Path]::Combine($app.cmdDir, "logs")
  New-Item -Force -ItemType Directory (Split-Path -Parent $logDir) > $null
  Start-Transcript ([System.IO.Path]::Combine($logDir, "$($app.cmdName)_$($app.now).log"))

  # exit code.
  $app.Add("cnst", @{
      SUCCESS = 0
      ERROR   = 1
    })

  # Init result
  $app.Add("result", $app.cnst.ERROR)

  # Init args.
  $app.Add("path", $path)

  # config.
  if ([string]::IsNullOrEmpty($app.path)) {
    $app.path = [System.IO.Path]::Combine($app.cmdDir, "$($app.cmdName).json")
  }
  $json = Get-Content -Encoding utf8 $app.path | ConvertFrom-Json
  $app.Add("cfg", $json)

  # excel enums
  $app.Add("ee", (Get-ExlConstEnums))

  log "[Start-Init] End" "Cyan"
}

<#
  .SYNOPSIS
    Get-ExlConstEnums
  .DESCRIPTION
    Get Excel const variables.
  .INPUTS
    - None
  .OUTPUTS
    - const values
#>
function Get-ExlConstEnums {

  [CmdletBinding()]
  [OutputType([object])]
  param()

  try {

    # Load excel enums if exist.
    $tmpenums = [System.IO.Path]::Combine($env:tmp, "EXCEL_ENUMS.json")

    if (Test-Path $tmpenums) {
      return Get-Content -Encoding utf8 $tmpenums | ConvertFrom-Json
    }

    # create Excel object
    $excel = New-Object -ComObject Excel.Application
    # create new Excel object
    $enums = [PSCustomObject]@{}
    # get all Excel exported types of type Enum
    $excel.GetType().Assembly.GetExportedTypes() | Where-Object { $_.IsEnum } | ForEach-Object {
      # create properties from enum values
      $enum = $_
      $enum.GetEnumNames() | ForEach-Object {
        $enums | Add-Member -MemberType NoteProperty -Name $_ -Value $enum::($_)
      }
    }
    $enums | ConvertTo-Json | Out-File -Encoding utf8 $tmpenums
    $enums
  }
  catch {
    Write-Host "[Get-ExlConstEnums] Error $_"
    throw $_
  }
  finally {
    if ($excel) { $excel.Quit() }
  }
}

<#
  .SYNOPSIS
    Main
  .DESCRIPTION
    Execute main
  .INPUTS
    - None
  .OUTPUTS
    - Result - 0 (SUCCESS), 1 (ERROR)
#>
function Start-Main {

  [CmdletBinding()]
  [OutputType([int])]
  param()

  try {
    log "[Start-Main] Start" "Cyan"
    $startTime = Get-Date

    Start-Init

    $EXCEL = New-Object -ComObject Excel.Application
    $EXCEL.Visible = $false
    $EXCEL.Application.ScreenUpdating = $false
    $EXCEL.Application.DisplayAlerts = $false
    $EXCEL.Application.EnableEvents = $false

    # Dst excel.
    $dstBook = $EXCEL.Workbooks.Add()

    $EXCEL.Application.Calculation = $app.ee.xlCalculationManual

    $dstSheet = $dstBook.Worksheets.Item(1)
    $dstBook.Worksheets | Where-Object {
      $_.Name -ne $dstSheet.Name
    } | ForEach-Object {
      [void]$_.Delete()
    }

    $i = 0
    $keys = New-Object System.Collections.ArrayList
    $cols = New-Object System.Collections.ArrayList
    $app.cfg.src | ForEach-Object {
      $target = $_
      $i++

      if ($target.before) {
        log "==================== [${i}] before exec ===================="
        log "Invoke-Expression $($target.before)" "Magenta"
        Invoke-Expression $target.before
      }

      log "==================== [${i}] book ===================="
      log ("src: [{0}]" -f $target.src) "Magenta"
      $src = [System.IO.Path]::GetFullPath($target.src)
      log ("src full path: [{0}]" -f $src) "Magenta"
      $BOOK = $EXCEL.Workbooks.Open($src)

      $target.sheets | ForEach-Object {
        log "-------------------- [${i}] sheet --------------------"
        $s = $_
        $SHEET = $BOOK.Worksheets.Item($s.name.src)
        $MAX_ROW = $SHEET.Cells.Item($SHEET.Rows.Count, $s.max.col).End($app.ee.xlUp).Row
        $MAX_COL = $SHEET.Cells.Item($s.max.row, $SHEET.Columns.Count).End($app.ee.xlToLeft).Column
        log ("MAX_ROW: {0}" -f $MAX_ROW) "Yellow"
        log ("MAX_COL: {0}" -f $MAX_COL) "Yellow"


        log "-------------------- [${i}] add columns --------------------"
        if ($s.addcols) {
          $s.addcols | ForEach-Object {
            $addcol = $_
            log "Insert column $($addcol.range[1])"
            [void]$SHEET.Columns($addcol.range[1]).Insert()
          }
          $MAX_COL = $SHEET.Cells.Item($s.max.row, $SHEET.Columns.Count).End($app.ee.xlToLeft).Column
          log ("MAX_COL: {0}" -f $MAX_COL) "Yellow"
          $s.addcols | ForEach-Object {
            $addcol = $_
            $SHEET.Cells.Item($addcol.range[0], $addcol.range[1]) = $addcol.header
            $formula = $ExecutionContext.InvokeCommand.ExpandString($addcol.formula)
            log ("formula: {0}" -f $formula) "Yellow"
            if ($addcol.r1c1) {
              $SHEET.Range($SHEET.Cells.Item($addcol.range[0] + 1, $addcol.range[1]), $SHEET.Cells.Item($MAX_ROW, $addcol.range[1])).FormulaR1C1 = $formula
            }
            else {
              $SHEET.Range($SHEET.Cells.Item($addcol.range[0] + 1, $addcol.range[1]), $SHEET.Cells.Item($MAX_ROW, $addcol.range[1])).Formula = $formula
            }
          }
        }

        if ($s.columns.key) {
          log "-------------------- [${i}] get keys --------------------"
          $r1 = [int]$ExecutionContext.InvokeCommand.ExpandString($s.columns.key.range[0][0])
          $r2 = [int]$ExecutionContext.InvokeCommand.ExpandString($s.columns.key.range[0][1])
          $r3 = [int]$ExecutionContext.InvokeCommand.ExpandString($s.columns.key.range[1][0])
          $r4 = [int]$ExecutionContext.InvokeCommand.ExpandString($s.columns.key.range[1][1])
          log ("csv range (after): [[{0}, {1}], [{2}, {3}]]" -f $r1, $r2, $r3, $r4) "Blue"
          $keyList = $SHEET.Range($SHEET.Cells.Item($r1, $r2), $SHEET.Cells.Item($r3, $r4)).Value2
          if ($s.columns.key.remove) {
            $keyList = $keyList | Where-Object { !$s.columns.key.remove.Contains($_) }
          }
          $keyList | ForEach-Object {
            [void]$keys.Add($_)
          }
        }

        $s.columns.dst | ForEach-Object {
          [void]$cols.Add($_)
        }

        log "-------------------- [${i}] worksheet copy --------------------"
        [void]$SHEET.Copy([System.Reflection.Missing]::Value, $dstBook.Worksheets.Item($i))
        $dstBook.Worksheets.Item($i + 1).Name = $s.name.dst

        if ($s.sort) {
          log "-------------------- [${i}] sort --------------------"
          $sh = $dstBook.Worksheets.Item($i + 1)
          if ($sh.FilterMode) {
            $sh.ShowAllData()
          }
          [void]$sh.Cells.Item(1, 1).AutoFilter()
          [void]$sh.AutoFilter.Sort.SortFields.Clear()
          [void]$sh.AutoFilter.Sort.SortFields.Add($sh.Cells.Item(1, $s.sort.col), [System.Reflection.Missing]::Value, $app.ee.($s.sort.direction))
          $sh.AutoFilter.Sort.Header = $app.ee.xlYes
          [void]$sh.AutoFilter.Sort.Apply()
        }

        if ($s.value) {
          log "-------------------- [${i}] value --------------------"
          $sh = $dstBook.Worksheets.Item($i + 1)
          if ($sh.FilterMode) {
            $sh.ShowAllData()
          }
          [void]$sh.Cells.Copy()
          [void]$sh.Cells.PasteSpecial($app.ee.xlPasteValues)
        }
      }

      if ($target.remove) {
        log ("Remove: {0}" -f $src) "DarkYellow"
        Remove-Item -Force $src
      }
    }

    $keys = $keys.ToArray() | Sort-Object | Get-Unique

    while ($true) {
      try {
        $dstSheet.Activate()
        $keys -join "`r`n" | Set-Clipboard
        [void]$dstSheet.Range($dstSheet.Cells.Item(2, 1), $dstSheet.Cells.Item($keys.Count + 1, 1)).PasteSpecial()
      }
      catch {
        log $_ "Yellow"
        log "Go to loop"
        continue
      }
      break
    }

    log "dst sheet name: $($app.cfg.dst.sheet.name)"
    $dstSheet.Name = $app.cfg.dst.sheet.name
    log "dst sheet key header: $($app.cfg.dst.key.header)"
    $dstSheet.Cells.Item(1, 1) = $app.cfg.dst.key.header
    $MAX_ROW = $dstSheet.Cells.Item($dstSheet.Rows.Count, $s.max.col).End($app.ee.xlUp).Row

    $cols | ForEach-Object {
      $col = $_
      log "Column no.$($col.col) header: $($col.header)" "Cyan"
      $dstSheet.Cells.Item(1, $col.col) = $col.header
      log ("Column formula: {0}" -f $col.formula) "Yellow"
      if ($col.r1c1) {
        $dstSheet.Range($dstSheet.Cells.Item(2, $col.col), $dstSheet.Cells.Item($keys.Count + 1, $col.col)).FormulaR1C1 = $col.formula
      }
      else {
        $dstSheet.Range($dstSheet.Cells.Item(2, $col.col), $dstSheet.Cells.Item($keys.Count + 1, $col.col)).Formulwa = $col.formula
      }
    }

    if ($app.cfg.dst.addcols) {
      log "-------------------- dst add columns --------------------"
      $app.cfg.dst.addcols | ForEach-Object {
        $addcol = $_
        log "Insert column $($addcol.range[1])"
        [void]$dstSheet.Columns($addcol.range[1]).Insert()
      }
      $MAX_COL = $dstSheet.Cells.Item($s.max.row, $dstSheet.Columns.Count).End($app.ee.xlToLeft).Column
      log ("MAX_COL: {0}" -f $MAX_COL) "Yellow"
      $app.cfg.dst.addcols | ForEach-Object {
        $addcol = $_
        $dstSheet.Cells.Item($addcol.range[0], $addcol.range[1]) = $addcol.header
        $formula = $ExecutionContext.InvokeCommand.ExpandString($addcol.formula)
        log ("formula: {0}" -f $formula) "Yellow"
        if ($addcol.r1c1) {
          $dstSheet.Range($dstSheet.Cells.Item($addcol.range[0] + 1, $addcol.range[1]), $dstSheet.Cells.Item($MAX_ROW, $addcol.range[1])).FormulaR1C1 = $formula
        }
        else {
          $dstSheet.Range($dstSheet.Cells.Item($addcol.range[0] + 1, $addcol.range[1]), $dstSheet.Cells.Item($MAX_ROW, $addcol.range[1])).Formula = $formula
        }
      }
    }

    if ($app.cfg.dst.formats) {
      log "-------------------- dst format --------------------"
      $app.cfg.dst.formats | ForEach-Object {
        $format = $_
        $r1 = [int]$ExecutionContext.InvokeCommand.ExpandString($format.range[0][0])
        $r2 = [int]$ExecutionContext.InvokeCommand.ExpandString($format.range[0][1])
        $r3 = [int]$ExecutionContext.InvokeCommand.ExpandString($format.range[1][0])
        $r4 = [int]$ExecutionContext.InvokeCommand.ExpandString($format.range[1][1])
        log ("Format range (after): [[{0}, {1}], [{2}, {3}]]" -f $r1, $r2, $r3, $r4) "Blue"
        $range = $dstSheet.Range($dstSheet.Cells.Item($r1, $r2), $dstSheet.Cells.Item($r3, $r4))
        log ("format: {0}" -f $format.format) "Green"
        $range.NumberFormatLocal = $format.format
      }
    }

    # windows freeze
    if ($app.cfg.dst.freeze) {
      log ("Windows freeze [{0}, {1}]" -f $app.cfg.dst.freeze[0], $app.cfg.dst.freeze[1]) "Green"
      [void]$dstSheet.Cells.Item($app.cfg.dst.freeze[0], $app.cfg.dst.freeze[1]).Select()
      $EXCEL.ActiveWindow.FreezePanes = $true
    }

    log "Entire column auto fit"
    [void]$dstSheet.Cells.EntireColumn.AutoFit()

    if ($app.cfg.dst.hide) {
      log "-------------------- dst hide --------------------"
      $app.cfg.dst.hide | ForEach-Object {
        log ("hide col: {0}" -f $_) "Blue"
        $dstSheet.Columns.Item($_).Hidden = $true
      }
    }

    if ($app.cfg.dst.group) {
      log "-------------------- dst group --------------------"
      $app.cfg.dst.group | ForEach-Object {
        log ("group col: {0}" -f $_) "Blue"
        [void]$dstSheet.Range($dstSheet.Columns($_[0]), $dstSheet.Columns($_[1])).Group()
      }
      [void]$dstSheet.Outline.ShowLevels([System.Reflection.Missing]::Value, 1)
    }

    if ($app.cfg.dst.sort) {
      log "-------------------- [${i}] dst sort --------------------"
      if ($dstSheet.FilterMode) {
        $dstSheet.ShowAllData()
      }
      [void]$dstSheet.Cells.Item($s.max.row, $s.max.col).AutoFilter()
      [void]$dstSheet.AutoFilter.Sort.SortFields.Clear()
      [void]$dstSheet.AutoFilter.Sort.SortFields.Add($dstSheet.Cells.Item($s.max.row, $app.cfg.dst.sort.col), [System.Reflection.Missing]::Value, $app.ee.($app.cfg.dst.sort.direction))
      $dstSheet.AutoFilter.Sort.Header = $app.ee.xlYes
      [void]$dstSheet.AutoFilter.Sort.Apply()
    }

    if ($app.cfg.dst.filter) {
      log "-------------------- [${i}] dst filter --------------------"
      $app.cfg.dst.filter | ForEach-Object {
        $term = $_
        if ($dstSheet.FilterMode) {
          $dstSheet.ShowAllData()
        }
        $criteria1 = & {
          if ($term.criteria1.GetType().BaseType.Name -eq "Array") {
            $term.criteria1 | ForEach-Object {
              $ExecutionContext.InvokeCommand.ExpandString($_)
            }
          }
          else {
            $ExecutionContext.InvokeCommand.ExpandString($term.criteria1)
          }
        }
        if ($criteria1 -match "^@") {
          $criteria1 = Invoke-Expression $criteria1
        }
        if ([string]::IsNullOrEmpty($term.operator)) {
          log ("--- field: {0}, criteria1: {1} ---" -f $term.field, $criteria1.ToString()) "Cyan"
          $dstSheet.Cells.Item($MAX_ROW, 1).AutoFilter($term.field, $criteria1) > $null
        }
        elseif ([string]::IsNullOrEmpty($term.criteria2)) {
          log ("--- field: {0}, criteria1: {1}, operator: {2} ({3}) ---" -f $term.field, $criteria1.ToString(), $term.operator, $app.ee.($term.operator)) "Cyan"
          $dstSheet.Cells.Item($MAX_ROW, 1).AutoFilter($term.field, $criteria1, $app.ee.($term.operator)) > $null
        }
        else {
          $criteria2 = $ExecutionContext.InvokeCommand.ExpandString($term.criteria2)
          if ($criteria2 -match "^@") {
            $criteria2 = Invoke-Expression $criteria2
          }
          log ("--- field: {0}, criteria1: {1}, operator: {2} ({3}), criteria2: {4} ---" -f $term.field, $criteria1.ToString(), $term.operator, $app.ee.($term.operator), $criteria2) "Cyan"
          $dstSheet.Cells.Item($MAX_ROW, 1).AutoFilter($term.field, $criteria1, $app.ee.($term.operator), $criteria2) > $null
        }
      }
    }

    $dst = [System.IO.Path]::GetFullPath($ExecutionContext.InvokeCommand.ExpandString($app.cfg.dst.dst))
    log ("dst full path: [{0}]" -f $dst) "Magenta"

    $EXCEL.Application.EnableEvents = $true
    $EXCEL.Application.Calculation = $app.ee.xlCalculationAutomatic

    New-Item -Force -ItemType Directory (Split-Path -Parent $dst) > $null
    $dstBook.SaveAs($dst)

    $app.result = $app.cnst.SUCCESS
  }
  catch {
    log "Error ! $_" "Red"
    Enable-RunspaceDebug -BreakAll
  }
  finally {
    if ($excel) { $excel.Quit() }
    $endTime = Get-Date
    $span = $endTime - $startTime
    log ("Elapsed time: {0} {1:00}:{2:00}:{3:00}.{4:000}" -f $span.Days, $span.Hours, $span.Minutes, $span.Seconds, $span.Milliseconds)
    log "[Start-Main] End" "Cyan"
    Stop-Transcript
  }
}

# Call main.
Start-Main
exit $app.result
