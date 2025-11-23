<#
  .SYNOPSIS
    exifrename.ps1
  .DESCRIPTION
    exiftool の結果を元にリネーム (移動) する
  .Last Change : 2025/11/23 13:32:11.
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
  Start-Transcript $app.logFile | Out-Null

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

      # Set output encoding to UTF-8 to handle special characters from exiftool
      $originalOutputEncoding = [System.Console]::OutputEncoding
      [System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8
      $exifJson = exiftool -json -d "%Y-%m-%d %H:%M:%S.%f" $file.FullName | ConvertFrom-Json
      [System.Console]::OutputEncoding = $originalOutputEncoding
      $dateStr = $exifJson.DateTimeOriginal
      if ([string]::IsNullOrEmpty($dateStr)) {
        $dateStr = $exifJson.CreateDate
      }

      if ([string]::IsNullOrEmpty($dateStr)) {
        log "Date is not found in $($file.FullName)" "Yellow"
        return
      }

      $date = Get-Date $dateStr

      # Add milliseconds if SubSecTimeOriginal is available
      if ($exifJson.SubSecTimeOriginal) {
        $millisecondsStr = $exifJson.SubSecTimeOriginal.ToString().Trim()
        if ($millisecondsStr -match "^\d+$") {
          # Pad with zeros to make it 3 digits, or truncate if longer
          $millisecondsStr = $millisecondsStr.PadRight(3, '0').Substring(0, 3)
          [int]$milliseconds = 0
          if ([int]::TryParse($millisecondsStr, [ref]$milliseconds)) {
            # DateTime objects are immutable, so we create a new one with updated milliseconds.
            $date = New-Object DateTime($date.Year, $date.Month, $date.Day, $date.Hour, $date.Minute, $date.Second, $milliseconds)
          }
        }
      }
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

      # Handle file collisions with sequential numbering
      if (Test-Path $destPath) {
        if ($app.cfg.sequential_format) {
          $index = 1
          $dirName = [System.IO.Path]::GetDirectoryName($destPath)
          $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($destPath)
          $extension = [System.IO.Path]::GetExtension($destPath)
          while ($true) {
            $sequentialPart = $app.cfg.sequential_format.Replace("{index}", $index)
            $newFileName = "$fileNameWithoutExt$sequentialPart$extension"
            $newDestPath = [System.IO.Path]::Combine($dirName, $newFileName)
            if (!(Test-Path $newDestPath)) {
              $destPath = $newDestPath
              break
            }
            $index++
          }
        } else {
          log "Destination file exists and sequential_format is not configured. Skipping: $($file.FullName)" "Yellow"
          return # 'return' will skip to the next item in ForEach-Object
        }
      }

      $destDir = [System.IO.Path]::GetDirectoryName($destPath)
      if (!(Test-Path $destDir)) {
        New-Item -ItemType Directory -Force -Path $destDir | Out-Null
      }

      log "Move $($file.FullName) to $destPath"
      Move-Item -Path $file.FullName -Destination $destPath
    }

    return $app.cnst.SUCCESS
  } catch {
    log "Error ! $($_ | Out-String)" "Red"
    return $app.cnst.ERROR
  } finally {
    log "[Start-Main] End"
    if ($app.logFile) {
      Stop-Transcript | Out-Null
    }
  }
}

# Call main.
exit Start-Main
