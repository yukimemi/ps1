<#
  .SYNOPSIS
    csv をマージする
  .DESCRIPTION
    指定したフォルダから csv ファイルを検索して
    指定されたキーでマージする
  .INPUTS
    - path: config path
  .OUTPUTS
    - 0: SUCCESS / 1: ERROR
  .Last Change : 2020/10/07 10:00:23.
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

  log "[Start-Init] End" "Cyan"
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

    $app.cfg.src | ForEach-Object {
      log ("src: [{0}]" -f $_) "Magenta"
      if (!(Test-Path $_)) {
        log ("src: [{0}] is not found ! skip !" -f $_) "DarkYellow"
        return
      }
      Get-ChildItem $_ | Where-Object {
        $_.FullName -match ($app.cfg.top -join "|")
      } | ForEach-Object {
        log ("top: [{0}]" -f $_.FullName)
        Get-ChildItem -Recurse $_.FullName | Where-Object {
          $_.FullName -match ($app.cfg.search -join "|")
        } | Where-Object {
          $_.Extension -eq ".csv"
        }
      }
    } | Set-Variable csvs

    $cols = $csvs | ForEach-Object { $_.FullName -replace $app.cfg.replace[0], $app.cfg.replace[1] } | Sort-Object | Get-Unique
    $keys = @{}

    $csvs | ForEach-Object {
      $csv = $_
      log ("csv: [{0}]" -f $csv.FullName) "Green"
      & {
        if ([string]::IsNullOrEmpty($app.cfg.header)) {
          Import-Csv -Encoding Default $csv.FullName
        }
        else {
          $csvData = Get-Content -Encoding Default $csv.FullName
          if ($app.cfg.headerreplace) {
            $csvData = Get-Content -Encoding Default $csv.FullName | Select-Object -Skip 1
            $csvData = @(($app.cfg.header -join ",")) + $csvData
          }
          $csvData | ConvertFrom-Csv
        }
      } | ForEach-Object {
        $record = $_
        $KEY = $record.($app.cfg.key)
        if (![string]::IsNullOrEmpty($app.cfg.eval.key)) {
          Invoke-Expression $app.cfg.eval.key
        }
        if ([string]::IsNullOrEmpty($KEY)) {
          log "[$($app.cfg.key)] is not found !" "Red"
          return
        }

        if (!$keys.ContainsKey($KEY)) {
          log "New ! $($KEY) [$($csv.FullName)]" "DarkMagenta"
          $keys.$KEY = [PSCustomObject]@{
            $app.cfg.key = $KEY
          }

          $cols | ForEach-Object {
            $keys.$KEY | Add-Member $_ $app.cfg.init
          }
        }

        $VALUE = $record.($app.cfg.value)
        if (![string]::IsNullOrEmpty($app.cfg.eval.value)) {
          Invoke-Expression $app.cfg.eval.value
        }
        $keys.$KEY.($csv.FullName -replace $app.cfg.replace[0], $app.cfg.replace[1]) = $VALUE
      }
      if ($app.cfg.remove) {
        log ("Remove: {0}" -f $csv.FullName) "DarkYellow"
        Remove-Item -Force $csv.FullName
      }
    }

    log "Export to $($app.cfg.dst)" "Green"
    New-Item -Force -ItemType Directory (Split-Path -Parent $app.cfg.dst) | Out-Null
    $keys.GetEnumerator() | ForEach-Object {
      $_.Value
    } | Set-Variable out 
    if ($app.cfg.filter) {
      $out = $out | Where-Object {
        Invoke-Expression $app.cfg.filter
      }
    }
    if ($app.cfg.group) {
      Write-Host "stop"
    }
    if ($app.cfg.sort) {
      $out = $out | Sort-Object $app.cfg.sort
    }
    $out | Export-Csv -NoTypeInformation -Encoding Default $app.cfg.dst

    $app.result = $app.cnst.SUCCESS

  }
  catch {
    log "Error ! $_" "Red"
  }
  finally {
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

# vim:fdm=syntax expandtab fdc=3 ft=ps1 ts=2 sw=2 sts=2 fenc=utf8 ff=dos:
