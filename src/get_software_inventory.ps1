#
# Get-SoftwareInventory.ps1
#
# GPOスタートアップスクリプトとして利用し、端末にインストールされている全ソフトウェア情報を取得します。
# 従来のWin32アプリ（レジストリ）に加え、ストアアプリ(Appx)の情報も網羅します。
#

$TargetBaseDir = "\\10.2.1.249\e$\SoftwareInventory"
$Today = Get-Date -Format "yyyy-MM-dd"
$HistoryDir = Join-Path $TargetBaseDir $Today
$LatestDir = Join-Path $TargetBaseDir "Latest"
$ComputerName = $env:COMPUTERNAME
$HistoryFile = Join-Path $HistoryDir "$ComputerName.csv"
$LatestFile = Join-Path $LatestDir "$ComputerName.csv"

# 負荷分散ランダムウェイト
0..3600 | Get-Random | Start-Sleep

# 出力先ディレクトリの準備
@($HistoryDir, $LatestDir) | ForEach-Object {
  if (-not (Test-Path $_)) {
    New-Item -Path $_ -ItemType Directory -Force | Out-Null
  }
}

$SoftwareList = New-Object System.Collections.Generic.List[PSObject]

# --- 1. マシン全体(HKLM)のスキャン ---
$MachinePaths = @(
  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($Path in $MachinePaths) {
  Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
    if ([string]::IsNullOrWhiteSpace($_.DisplayName)) {
      return
    }
    $SoftwareList.Add([PSCustomObject]@{
        ComputerName   = $ComputerName
        Username       = "SYSTEM (Machine-wide)"
        DisplayName    = $_.DisplayName
        DisplayVersion = $_.DisplayVersion
        Publisher      = $_.Publisher
        InstallDate    = $_.InstallDate
        InstallSource  = "Machine"
        RegistryPath   = $_.PSPath
      })
  }
}

# --- 2. ユーザー個別(HKCU)のスキャン（レジストリロード） ---
$Profiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\S-1-5-21-*" -ErrorAction SilentlyContinue

foreach ($Profile in $Profiles) {
  $Sid = $Profile.PSChildName
  $ProfilePath = $Profile.ProfileImagePath
  $Username = Split-Path $ProfilePath -Leaf
  $NtUserDat = Join-Path $ProfilePath "NTUSER.DAT"

  if (-not (Test-Path $NtUserDat)) {
    continue
  }

  $LoadedTempKey = $null
  $UserRegistryBase = "Registry::HKEY_USERS\$Sid"

  if (-not (Test-Path $UserRegistryBase)) {
    $TempKeyName = "TempHive_$Sid"
    $UserRegistryBase = "Registry::HKEY_USERS\$TempKeyName"
    try {
      reg load "HKU\$TempKeyName" "$NtUserDat" 2>&1 | Out-Null
      $LoadedTempKey = $TempKeyName
    } catch {
      continue
    }
  }

  $UserUninstallPaths = @(
    "$UserRegistryBase\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "$UserRegistryBase\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
  )

  foreach ($Path in $UserUninstallPaths) {
    Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
      if ([string]::IsNullOrWhiteSpace($_.DisplayName)) {
        return
      }
      $SoftwareList.Add([PSCustomObject]@{
          ComputerName   = $ComputerName
          Username       = $Username
          DisplayName    = $_.DisplayName
          DisplayVersion = $_.DisplayVersion
          Publisher      = $_.Publisher
          InstallDate    = $_.InstallDate
          InstallSource  = "User"
          RegistryPath   = $_.PSPath
        })
    }
  }

  if ($LoadedTempKey) {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    reg unload "HKU\$LoadedTempKey" 2>&1 | Out-Null
  }
}

# --- 3. ストアアプリ(Appx)のスキャン ---
# 全ユーザー対象のパッケージを取得
try {
  Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | ForEach-Object {
    # システムコンポーネントすぎるものは除外したい場合はフィルタが必要ですが、
    # VAIO設定などのツールを拾うために一旦すべて取得します。
    $SoftwareList.Add([PSCustomObject]@{
        ComputerName   = $ComputerName
        Username       = "All Users (Appx)"
        DisplayName    = $_.Name
        DisplayVersion = $_.Version
        Publisher      = $_.Publisher
        InstallDate    = "" # Appxからは取得が難しいため空
        InstallSource  = "Appx"
        RegistryPath   = $_.PackageFullName
      })
  }
} catch {
}

# --- 4. CSV出力 ---
if ($SoftwareList.Count -gt 0) {
  $SortedList = $SoftwareList | Sort-Object DisplayName
  $SortedList | Export-Csv -Path $HistoryFile -NoTypeInformation -Encoding utf8
  $SortedList | Export-Csv -Path $LatestFile -NoTypeInformation -Encoding utf8
}

