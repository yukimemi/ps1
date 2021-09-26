<#
  .SYNOPSIS
    csvupdate
  .DESCRIPTION
    csv で key に従って重複したら設定に従って update する
  .INPUTS
    - path (cfg path)
  .OUTPUTS
    - 0: SUCCESS / 1: ERROR
  .Last Change : 2021/08/21 18:54:23.
#>
$a = $args
$path = & { if ([string]::IsNullOrEmpty($a[0])) { "" } else { $a[0] } }
$master = & { if ([string]::IsNullOrEmpty($a[1])) { "" } else { $a[1] } }
$update = & { if ([string]::IsNullOrEmpty($a[2])) { "" } else { $a[2] } }
$merge = & { if ([string]::IsNullOrEmpty($a[3])) { "" } else { $a[3] } }

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

    if ([string]::IsNullOrEmpty($master)) {
      $master = $ExecutionContext.InvokeCommand.ExpandString($app.cfg.master)
    }
    if ([string]::IsNullOrEmpty($update)) {
      $update = $ExecutionContext.InvokeCommand.ExpandString($app.cfg.update)
    }
    if ([string]::IsNullOrEmpty($merge)) {
      $merge = $ExecutionContext.InvokeCommand.ExpandString($app.cfg.merge)
    }

    if (!(Test-Path -PathType Leaf $master)) {
      $master = $update
    }
    if (!(Test-Path -PathType Leaf $update)) {
      $update = $master
    }
    $master = Import-Csv $master
    $update = Import-Csv $update

    $keyId = "__key"
    $addKey = {
      param([object]$data, [string[]]$keys)
      $data | ForEach-Object {
        $d = $_
        $key = ($keys | ForEach-Object { $d.$_ }) -join "_"
        $d | Add-Member -Force -MemberType NoteProperty $keyId $key
        $d
      }
    }

    # Add key column.
    $master = & $addKey $master $app.cfg.key
    $update = & $addKey $update $app.cfg.key

    $row = 0
    $masterClone = $master.Clone()
    $masterClone | ForEach-Object {
      $row++
      $m = $_
      log "[${row}] master key: [$($m.$keyId)]" "Magenta"
      $u = $null
      $u = $update | Where-Object { $m.$keyId -eq $_.$keyId } | Select-Object -First 1
      if ($null -ne $u) {
        $app.cfg.updates | ForEach-Object {
          # log "col: [$($_.col)], up: [$($_.up)]" "Cyan"
          $up = Invoke-Expression $_.up
          if ($up) {
            log "up ! $($_.col): [$($m.($_.col))] → [$($u.($_.col))]" "Green"
            $m.($_.col) = $u.($_.col)
          }
        }
      }
      $m
    } | Set-Variable updateMaster

    $update | Where-Object {
      $u = $_
      $m = $null
      $m = $master | Where-Object { $u.$keyId -eq $_.$keyId } | Select-Object -First 1
      if ($null -eq $m) {
        log "update only: [$($u.$keyId)]" "Green"
        $u
      }
    } | Set-Variable updateOnly

    if ($null -ne $updateOnly) {
      $updateMaster += $updateOnly
    }

    log "Output update file: [${merge}]" "Magenta"
    $updateMaster | Export-Csv -NoTypeInformation -Encoding utf8 $merge

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
