<#
  .SYNOPSIS
    urldec
  .DESCRIPTION
    urldecodeする
  .INPUTS
    - None
  .OUTPUTS
    - 0: SUCCESS / 1: ERROR
  .Last Change : 2021/09/26 15:28:41.
#>
param()

$ErrorActionPreference = "Stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire
# Enable-RunspaceDebug -BreakAll
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Web

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
  trap { Write-Host "[Start-Init] Error $_"; throw $_ }

  Write-Host "[Start-Init] Start"

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

  # const value.
  $app.Add("cnst", @{
      SUCCESS = 0
      ERROR   = 1
    })

  # Init result
  $app.Add("result", $app.cnst.ERROR)

  Write-Host "[Start-Init] End"
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
    # StartTime
    $startTime = Get-Date

    Start-Init

    Get-Clipboard | ForEach-Object {
      $urldec = [System.Web.HttpUtility]::UrlDecode($_)
      Write-Host $urldec
      $urldec
    } | Set-Clipboard

    $app.result = $app.cnst.SUCCESS

  }
  catch {
    Write-Host "Error ! $_"
  }
  finally {
    $endTime = Get-Date
    $span = $endTime - $startTime
    Write-Host ("Elapsed time: {0} {1:00}:{2:00}:{3:00}.{4:000}" -f $span.Days, $span.Hours, $span.Minutes, $span.Seconds, $span.Milliseconds)
    $app.result
  }
}

# Call main.
exit Start-Main

