<#
  .SYNOPSIS
    mht2jpg
  .DESCRIPTION
    psr で採取した情報を jpg に変換する
  .INPUTS
    - path: psr zip path
  .OUTPUTS
    - 0: SUCCESS / 1: ERROR
  .Last Change : 2021/04/28 16:45:18.
#>
param(
  [Parameter()]
  [string]$path = (Read-Host "psr zip path")
)

$ErrorActionPreference = "Stop"
$path = $path -replace '"', ""
$dir = [System.IO.Path]::GetDirectoryName($path)
$dstName = [System.IO.Path]::GetFileNameWithoutExtension($path)
$dst = [System.IO.Path]::Combine($dir, $dstName)
New-Item -Force -ItemType Directory $dst > $null
Write-Host "[${path}] -> [${dst}]"
Expand-Archive -Force -LiteralPath $path $dst

$path = Get-ChildItem -File $dst | Where-Object { $_.Extension -eq ".mht" } | Select-Object -ExpandProperty FullName
Write-Host "mht: ${path}"

$nameRE = [regex]"^Content-Location: (?<Name>screenshot\d{4}.JPEG)$"
$nextRE = [regex]"^--=_NextPart"

$imagePart = $false
$base64 = ""

switch -regex -file $path {
  '^$' {}

  $nameRE {
    $name = $nameRE.Match($_).Groups['Name'].Value
    $file = Join-Path $dst $name

    $imagePart = $true
  }

  $nextRE {
    if ($imagePart) {
      Write-Host "Save: [${file}]"
      $byte = [Convert]::FromBase64String($base64)
      [IO.File]::WriteAllBytes($file, $byte)

      $imagePart = $false
      $base64 = ""
    }
  }

  default {
    if ($imagePart) {
      $base64 += $_
    }
  }
}

