<#
.SYNOPSIS
  暗号化/復号化
.DESCRIPTION
  暗号化 (Encrypt-Plain)、復号化 (Decrypt-Secure) の2関数を提供する
.Last Change : 2018/07/29 12:19:17.
#>
$ErrorActionPreference = "Stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire

<#
.SYNOPSIS
  暗号化
.DESCRIPTION
  平文を暗号化してファイルに保存する
.EXAMPLE
  Encrypt-Plain -Plain "12345" -Path "C:\crypto\secret.enc"
  # 平文「12345」を暗号化して「C:\crypto\secret.enc」に保存します
#>
function Encrypt-Plain {

  [CmdletBinding()]
  [OutputType([void])]
  param(
    # 暗号化する平文を指定します
    [parameter(mandatory)]
    [string]$Plain,

    # 暗号化した平文の保存先を指定します
    [parameter(mandatory)]
    [string]$Path
  )
  trap { Write-Host "[Encrypt-Plain] Error $_"; throw $_ }

  # 暗号化用のバイト配列を準備
  $key = [byte[]]@(0x63, 0x72, 0x79, 0x70, 0x74, 0x6f, 0x65, 0x6e, 0x63, 0x64, 0x65, 0x63)
  $key += $key

  $secure = ConvertTo-SecureString -String $Plain -AsPlainText -Force
  $enc = ConvertFrom-SecureString -SecureString $secure -key $key

  # 書き込み
  New-Item -Force -ItemType Directory (Split-Path -Parent $Path) > $null
  $enc | Set-Content $Path
}

<#
.SYNOPSIS
  復号化
.DESCRIPTION
  ファイルから暗号化された文字を読んで復号化する
.OUTPUTS
  [string] 復号化した平文
.EXAMPLE
  $plain = Decrypt-Secure -Path "C:\crypto\secret.enc"
  # ファイル「C:\crypto\secret.enc」内の暗号化された文字を復号化して返します
#>
function Decrypt-Secure {

  [CmdletBinding()]
  [OutputType([string])]
  param(
    # 復号化する文字が含まれるファイルパスを指定します
    [parameter(mandatory)]
    [string]$Path
  )
  trap { Write-Host "[Decrypt-Secure] Error $_"; throw $_ }

  # 複号化用のバイト配列を準備
  $key = [byte[]]@(0x63, 0x72, 0x79, 0x70, 0x74, 0x6f, 0x65, 0x6e, 0x63, 0x64, 0x65, 0x63)
  $key += $key

  # 暗号化された標準文字列をインポートしてSecureStringに変換
  $secure = Get-Content $Path | ConvertTo-SecureString -key $key

  # 復号化
  $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
}

