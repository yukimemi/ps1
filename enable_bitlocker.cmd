@set __SCRIPTPATH=%~f0&@powershell -NoProfile -ExecutionPolicy ByPass -InputFormat None "$s=[scriptblock]::create((gc \"%~f0\"|?{$_.readcount -gt 2})-join\"`n\");&$s" %*
@exit /b %errorlevel%
<#
.SYNOPSIS
  BitLocker�� (�p�X���[�h)
.DESCRIPTION
  �p�X���[�h�����Ńh���C�u��BitLocker������
.Last Change : 2018/07/30 03:23:27.
#>
param(
  # �h���C�u
  [parameter(mandatory)]
  [string]$Drive,
  # �p�X���[�h
  [parameter(mandatory)]
  [string]$Pass
)
$ErrorActionPreference = "Stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire

trap { Write-Host "[enable_bitlocker] Error $_"; exit 1 }

$secure = ConvertTo-SecureString $Pass -AsPlainText -Force

Enable-BitLocker -MountPoint ${Drive}: -UsedSpaceOnly -EncryptionMethod AES128 -Password $secure -PasswordProtector -ErrorAction Stop

exit 0