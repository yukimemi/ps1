<#
.SYNOPSIS
  �Í���/������
.DESCRIPTION
  �Í��� (Encrypt-Plain)�A������ (Decrypt-Secure) ��2�֐���񋟂���
.Last Change : 2018/07/29 12:19:17.
#>
$ErrorActionPreference = "Stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire

<#
.SYNOPSIS
  �Í���
.DESCRIPTION
  �������Í������ăt�@�C���ɕۑ�����
.EXAMPLE
  Encrypt-Plain -Plain "12345" -Path "C:\crypto\secret.enc"
  # �����u12345�v���Í������āuC:\crypto\secret.enc�v�ɕۑ����܂�
#>
function Encrypt-Plain {

  [CmdletBinding()]
  [OutputType([void])]
  param(
    # �Í������镽�����w�肵�܂�
    [parameter(mandatory)]
    [string]$Plain,

    # �Í������������̕ۑ�����w�肵�܂�
    [parameter(mandatory)]
    [string]$Path
  )
  trap { Write-Host "[Encrypt-Plain] Error $_"; throw $_ }

  # �Í����p�̃o�C�g�z�������
  $key = [byte[]]@(0x63, 0x72, 0x79, 0x70, 0x74, 0x6f, 0x65, 0x6e, 0x63, 0x64, 0x65, 0x63)
  $key += $key

  $secure = ConvertTo-SecureString -String $Plain -AsPlainText -Force
  $enc = ConvertFrom-SecureString -SecureString $secure -key $key

  # ��������
  New-Item -Force -ItemType Directory (Split-Path -Parent $Path) > $null
  $enc | Set-Content $Path
}

<#
.SYNOPSIS
  ������
.DESCRIPTION
  �t�@�C������Í������ꂽ������ǂ�ŕ���������
.OUTPUTS
  [string] ��������������
.EXAMPLE
  $plain = Decrypt-Secure -Path "C:\crypto\secret.enc"
  # �t�@�C���uC:\crypto\secret.enc�v���̈Í������ꂽ�����𕜍������ĕԂ��܂�
#>
function Decrypt-Secure {

  [CmdletBinding()]
  [OutputType([string])]
  param(
    # ���������镶�����܂܂��t�@�C���p�X���w�肵�܂�
    [parameter(mandatory)]
    [string]$Path
  )
  trap { Write-Host "[Decrypt-Secure] Error $_"; throw $_ }

  # �������p�̃o�C�g�z�������
  $key = [byte[]]@(0x63, 0x72, 0x79, 0x70, 0x74, 0x6f, 0x65, 0x6e, 0x63, 0x64, 0x65, 0x63)
  $key += $key

  # �Í������ꂽ�W����������C���|�[�g����SecureString�ɕϊ�
  $secure = Get-Content $Path | ConvertTo-SecureString -key $key

  # ������
  $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
}

