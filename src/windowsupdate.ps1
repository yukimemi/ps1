<#
  .SYNOPSIS
    Windows Update
  .DESCRIPTION
    Execute Windows Update
  .INPUTS
    - reboot: false (default): No reboot computer.
            : true           : reboot computer if required.
  .OUTPUTS
    - 0: SUCCESS / 1: ERROR
  .Last Change : 2021/08/21 18:54:23.
#>
param(
  [Parameter()]
  [bool]$reboot = $false
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
  $app.Add("reboot", $reboot)

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

    # Search.
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")
    log "List of applicable items on the machine." "Green"
    if ($searchResult.Updates.Count -eq 0) {
      log "There are no applicable updates."
      $app.result = $app.cnst.SUCCESS
      return
    }

    # Download.
    $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    $isDownload = $false
    $searchResult.Updates | ForEach-Object {
      $update = $_
      log "Title: [$($update.Title)], LastDeploymentChangeTime: [$($update.LastDeploymentChangeTime)], MaxDownloadSize: [$($update.MaxDownloadSize)], IsDownloaded: [$($update.IsDownloaded)], Description: [$($update.Description)]" "Green"
      if (!$update.IsDownloaded) {
        [void]$updatesToDownload.Add($update)
        $isDownload = $true
      }
    }

    if ($isDownload) {
      log "Downloading updates..." "Magenta"
      $downloader = $updateSession.CreateUpdateDownloader()
      $downloader.Updates = $updatesToDownload
      $result = $downloader.Download()
      log $result
    }
    else {
      log "All updates are already downloaded."
    }

    # Install.
    $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    log "Creating collection of downloaded updates to install..." "Green"
    $searchResult.Updates | ForEach-Object {
      $update = $_
      log "Title: [$($update.Title)], LastDeploymentChangeTime: [$($update.LastDeploymentChangeTime)], MaxDownloadSize: [$($update.MaxDownloadSize)], IsDownloaded: [$($update.IsDownloaded)], Description: [$($update.Description)]" "Green"
      if ($update.IsDownloaded) {
        [void]$updatesToInstall.Add($update)
      }
    }

    if ($updatesToInstall.Count -eq 0 ) {
      log "Not ready for installation." "Yellow"
      $app.result = $app.cnst.SUCCESS
      return
    }

    log "Installing $($updatesToInstall.Count) updates..."
    $installer = $updateSession.CreateUpdateInstaller()
    $installer.Updates = $updatesToInstall
    $result = $installer.Install()
    if ($result.ResultCode -eq 2) {
      log "All updates installed successfully." "Green"
    }
    else {
      log "Some updates could not installed." "Yellow"
    }

    if ($result.RebootRequired) {
      log "One or more updates are requiring reboot." "Magenta"
      if ($app.reboot) {
        log "Reboot system now !!" "Red"
        shutdown.exe /r /t 0
      }
    }
    else {
      log "Finished. Reboot are not required."
    }

    $app.result = $app.cnst.SUCCESS
  }
  catch {
    log "Error ! $_" "Red"
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
