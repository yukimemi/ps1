<#
  .SYNOPSIS
    exifrename.ps1
  .DESCRIPTION
    exiftool の結果を元にリネーム (移動) する
  .Last Change : 2025/11/23 02:52:43.
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$in_dir,
  [Parameter(Mandatory = $true)]
  [string]$out_dir,
  [string]$cfg
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
  trap {
    Write-Host "[log] Error $_" "Red"; throw $_
  }

  $now = Get-Date -f "yyyy/MM/dd HH:mm:ss.fff"
  if ($color) {
    Write-Host -ForegroundColor $color "${now} ${msg}"
  } else {
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
  trap {
    log "[Start-Init] Error $_" "Red"; throw $_
  }

  log "[Start-Init] Start"

  $script:app = @{}

  $cmdFullPath = & {
    if ($env:__SCRIPTPATH) {
      return [System.IO.Path]::GetFullPath($env:__SCRIPTPATH)
    } else {
      return [System.IO.Path]::GetFullPath($script:MyInvocation.MyCommand.Path)
    }
  }
  $app.Add("cmdFile", $cmdFullPath)
  $app.Add("cmdDir", [System.IO.Path]::GetDirectoryName($app.cmdFile))
  $app.Add("cmdName", [System.IO.Path]::GetFileNameWithoutExtension($app.cmdFile))
  $app.Add("cmdFileName", [System.IO.Path]::GetFileName($app.cmdFile))

  $app.Add("pwd", [System.IO.Path]::GetFullPath((Get-Location).Path))

  # log
  $app.Add("now", (Get-Date -Format "yyyyMMddTHHmmssfffffff"))
  $app.Add("logDir", [System.IO.Path]::Combine($app.cmdDir, "logs"))
  $app.Add("logFile", [System.IO.Path]::Combine($app.logDir, "$($app.cmdName)_$($app.now).log"))
  $app.Add("logName", [System.IO.Path]::GetFileNameWithoutExtension($app.logFile))
  $app.Add("logFileName", [System.IO.Path]::GetFileName($app.logFile))
  New-Item -Force -ItemType Directory $app.logDir | Out-Null
  Start-Transcript $app.logFile

  # const value.
  $app.Add("cnst", @{
      SUCCESS = 0
      ERROR   = 1
    })

  # config.
  if ([string]::IsNullOrEmpty($cfg)) {
    $app.Add("cfgPath", [System.IO.Path]::Combine($app.cmdDir, "$($app.cmdName).json"))
  } else {
    $app.Add("cfgPath", $cfg)
  }
  if (!(Test-Path $app.cfgPath)) {
    log "$($app.cfgPath) is not found ! finish ..."
    throw "$($app.cfgPath) is not found !"
  }
  $json = Get-Content -Encoding utf8 $app.cfgPath | ConvertFrom-Json
  $app.Add("cfg", $json)

  # Init result
  $app.Add("result", $app.cnst.ERROR)

  log "[Start-Init] End"
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
    Start-Init
    log "[Start-Main] Start"

    if (!(Test-Path $in_dir)) {
      log "Input directory not found: $in_dir" "Red"
      throw "Input directory not found: $in_dir"
    }

    # Check exiftool
    try {
      exiftool -ver | Out-Null
    } catch {
      log "exiftool not found. Please install exiftool." "Red"
      throw "exiftool not found."
    }

    $gci_params = @{
      Path = $in_dir
      File = $true
    }
    if ($app.cfg.recursive) {
      $gci_params.Add("Recurse", $true)
    }

    Get-ChildItem @gci_params | ForEach-Object {
      $file = $_
      log "Processing $($file.FullName)"

      $exifJson = exiftool -json -d "%Y-%m-%d %H:%M:%S" $file.FullName | ConvertFrom-Json
      $dateStr = $exifJson.DateTimeOriginal
      if ([string]::IsNullOrEmpty($dateStr)) {
        $dateStr = $exifJson.CreateDate
      }

      if ([string]::IsNullOrEmpty($dateStr)) {
        log "Date is not found in $($file.FullName)" "Yellow"
        return
      }

      $date = Get-Date $dateStr
      $newName = $app.cfg.rename_format
      $newName = $newName.Replace("{yyyymmdd}", $date.ToString("yyyyMMdd"))
      $newName = $newName.Replace("{yyyy}", $date.ToString("yyyy"))
      $newName = $newName.Replace("{mm}", $date.ToString("MM"))
      $newName = $newName.Replace("{dd}", $date.ToString("dd"))
      $newName = $newName.Replace("{hhmissfff}", $date.ToString("HHmmssfff"))
      $newName = $newName.Replace("{hhmiss}", $date.ToString("HHmmss"))
      $newName = $newName.Replace("{hhmi}", $date.ToString("HHmm"))
      $newName = $newName.Replace("{hh}", $date.ToString("HH"))
      $newName = $newName.Replace("{mi}", $date.ToString("mm"))
      $newName = $newName.Replace("{ss}", $date.ToString("ss"))
      $newName = $newName.Replace("{fff}", $date.ToString("fff"))
      $newName = $newName.Replace("{filename}", $file.BaseName)
      $newName = $newName.Replace("{ext}", $file.Extension.Substring(1))

      $destPath = [System.IO.Path]::Combine($out_dir, $newName)

      $destDir = [System.IO.Path]::GetDirectoryName($destPath)
      if (!(Test-Path $destDir)) {
        New-Item -ItemType Directory -Force -Path $destDir | Out-Null
      }

      log "Move $($file.FullName) to $destPath"
      Move-Item -Path $file.FullName -Destination $destPath
    }

    $app.result = $app.cnst.SUCCESS
  } catch {
    log "Error ! $($_ | Out-String)" "Red"
    $app.result = $app.cnst.ERROR
  } finally {
    log "[Start-Main] End"
    if ($app.logFile) {
      Stop-Transcript
    }
  }
  return $app.result
}

# Call main.
$exitCode = Start-Main
exit $exitCode
