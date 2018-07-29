@set __SCRIPTPATH=%~f0&@powershell -NoProfile -ExecutionPolicy ByPass -InputFormat None "$s=[scriptblock]::create((gc \"%~f0\"|?{$_.readcount -gt 2})-join\"`n\");&$s" %*
@exit /b %errorlevel%
<#
.SYNOPSIS
  BitLocker 解除 (パスワード)
.DESCRIPTION
  パスワード方式でBitLocker を解除する
.Last Change : 2018/07/30 03:22:31.
#>
param(
  # ドライブ
  [parameter(mandatory)]
  [string]$Drive,
  # パスワード
  [parameter(mandatory)]
  [string]$Pass
)
$ErrorActionPreference = "Stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire

trap { Write-Host "[unlock_bitlocker] Error $_"; exit 1 }

$secure = ConvertTo-SecureString $Pass -AsPlainText -Force

Unlock-BitLocker -MountPoint ${Drive}: -Password $secure -ErrorAction Stop

exit 0
