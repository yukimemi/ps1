<#
.SYNOPSIS
  暗号化/復号
.DESCRIPTION
  暗号化 (Encrypt-Plain)、復号 (Decrypt-Secure) の2関数を提供する
.Last Change : 2018/11/09 17:15:06.
#>
param(
  # 暗号化モードで動作
  [switch]$Enc = $false,
  # 復号モードで動作 (デフォルト)
  [switch]$Dec = $true,
  # 暗号化する平文
  [string]$Plain,
  # 暗号化ファイルパス
  [string]$Path
)
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
  復号
.DESCRIPTION
  ファイルから暗号化された文字を読んで復号する
.OUTPUTS
  [string] 復号した平文
.EXAMPLE
  $plain = Decrypt-Secure -Path "C:\crypto\secret.enc"
  # ファイル「C:\crypto\secret.enc」内の暗号化された文字を復号して返します
#>
function Decrypt-Secure {

  [CmdletBinding()]
  [OutputType([string])]
  param(
    # 復号する文字が含まれるファイルパスを指定します
    [parameter(mandatory)]
    [string]$Path
  )
  trap { Write-Host "[Decrypt-Secure] Error $_"; throw $_ }

  # 復号用のバイト配列を準備
  $key = [byte[]]@(0x63, 0x72, 0x79, 0x70, 0x74, 0x6f, 0x65, 0x6e, 0x63, 0x64, 0x65, 0x63)
  $key += $key

  # 暗号化された標準文字列をインポートしてSecureStringに変換
  $secure = Get-Content $Path | ConvertTo-SecureString -key $key

  # 復号
  $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
}


# cmd 形式の時以外では処理終了
if (![string]::IsNullOrEmpty($script:MyInvocation.MyCommand.Path)) {
  exit 0
}

# Path は必須
if ([string]::IsNullOrEmpty($Path)) {
  Write-Host "-Path パラメータは必須です"
  exit 1
}

# cmd 形式で実行された場合
try {
  if ($Enc) {
    Encrypt-Plain -Plain $Plain -Path $Path
  } else {
    Decrypt-Secure -Path $Path
  }
  exit 0
} catch {
  Write-Host "[crypto.ps1] Error $_"
  exit 1
}

