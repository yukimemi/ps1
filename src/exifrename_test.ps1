#
# exifrename_test.ps1
#
# Last Change: 2025/11/23 00:55:57.
#

$ErrorActionPreference = "Stop"

$base_dir = (Get-Location).Path
$script_dir = Join-Path $base_dir "src"
$script_path = Join-Path $script_dir "exifrename.ps1"

$tests = @(
  @{
    Name = "Path with spaces"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in dir with space"
      $out_dir = Join-Path $base_dir "tmp\out dir with space"
      $cfg = Join-Path $script_dir "exifrename.json"
      $test_img_src = Join-Path $base_dir "tmp\test with space.jpg"
      $test_img = Join-Path $in_dir "test.jpg"
      $out_file = Join-Path $out_dir "20251122/20251122_123456789.jpg"

      try {
        # Create a dummy jpg file
        $bitmap = New-Object System.Drawing.Bitmap(1, 1)
        $bitmap.SetPixel(0, 0, [System.Drawing.Color]::Red)
        $bitmap.Save($test_img_src, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Dispose()
        exiftool "-DateTimeOriginal=2025:11:22 12:34:56" "-SubSecTimeOriginal=789" -overwrite_original $test_img_src | Out-Null

        New-Item -ItemType Directory -Force -Path $in_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $out_dir | Out-Null
        Copy-Item -Path $test_img_src -Destination $test_img

        & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $cfg
        if (!(Test-Path $out_file)) {
          throw "Test failed: Output file '$out_file' not found in path with spaces."
        }
      } finally {
        if (Test-Path $in_dir) {
          Remove-Item -Recurse -Force $in_dir
        }
        if (Test-Path $out_dir) {
          Remove-Item -Recurse -Force $out_dir
        }
        if (Test-Path $test_img_src) {
          Remove-Item -Force $test_img_src
        }
      }
    }
  },
  @{
    Name = "Happy Path with Milliseconds"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in"
      $out_dir = Join-Path $base_dir "tmp\out"
      $cfg = Join-Path $script_dir "exifrename.json"
      $test_img_src = Join-Path $base_dir "tmp\test.jpg"
      $test_img = Join-Path $in_dir "test.jpg"
      $out_file = Join-Path $out_dir "20251122/20251122_123456789.jpg"

      try {
        # Create a dummy jpg file
        $bitmap = New-Object System.Drawing.Bitmap(1, 1)
        $bitmap.SetPixel(0, 0, [System.Drawing.Color]::Red)
        $bitmap.Save($test_img_src, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Dispose()
        exiftool "-DateTimeOriginal=2025:11:22 12:34:56" "-SubSecTimeOriginal=789" -overwrite_original $test_img_src | Out-Null

        New-Item -ItemType Directory -Force -Path $in_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $out_dir | Out-Null
        Copy-Item -Path $test_img_src -Destination $test_img

        & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $cfg
        if (!(Test-Path $out_file)) {
          throw "Test failed: Output file '$out_file' not found."
        }
      } finally {
        if (Test-Path $in_dir) {
          Remove-Item -Recurse -Force $in_dir
        }
        if (Test-Path $out_dir) {
          Remove-Item -Recurse -Force $out_dir
        }
        if (Test-Path $test_img_src) {
          Remove-Item -Force $test_img_src
        }
      }
    }
  },
  @{
    Name = "No EXIF date"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in"
      $out_dir = Join-Path $base_dir "tmp\out"
      $cfg = Join-Path $script_dir "exifrename.json"
      $test_img_src = Join-Path $base_dir "tmp\test.jpg"
      $test_img = Join-Path $in_dir "test_no_exif.jpg"
      $out_file = Join-Path $out_dir "test_no_exif.jpg"

      try {
        # Create a dummy jpg file
        $bitmap = New-Object System.Drawing.Bitmap(1, 1)
        $bitmap.SetPixel(0, 0, [System.Drawing.Color]::Red)
        $bitmap.Save($test_img_src, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Dispose()

        New-Item -ItemType Directory -Force -Path $in_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $out_dir | Out-Null
        Copy-Item -Path $test_img_src -Destination $test_img
        # Remove exif data
        exiftool -all= -overwrite_original $test_img | Out-Null

        $result = & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $cfg
        if ($result) {
          throw "Test failed: Script should have returned false."
        }
        if (!(Test-Path $test_img)) {
          throw "Test failed: Input file should not be moved."
        }
      } finally {
        if (Test-Path $in_dir) {
          Remove-Item -Recurse -Force $in_dir
        }
        if (Test-Path $out_dir) {
          Remove-Item -Recurse -Force $out_dir
        }
        if (Test-Path $test_img_src) {
          Remove-Item -Force $test_img_src
        }
      }
    }
  },
  @{
    Name = "Input directory does not exist"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in"
      $out_dir = Join-Path $base_dir "tmp\out"
      $cfg = Join-Path $script_dir "exifrename.json"

      # Ensure the input directory does not exist
      if (Test-Path $in_dir) {
        Remove-Item -Recurse -Force $in_dir
      }

      & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $cfg | Out-Null
      
      if ($LASTEXITCODE -eq 0) {
        throw "Test failed: Script should have returned a non-zero exit code."
      }
    }
  },
  @{
    Name = "Input directory is empty"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in"
      $out_dir = Join-Path $base_dir "tmp\out"
      $cfg = Join-Path $script_dir "exifrename.json"

      try {
        New-Item -ItemType Directory -Force -Path $in_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $out_dir | Out-Null

        & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $cfg
        if ((Get-ChildItem -Path $out_dir).Count -ne 0) {
          throw "Test failed: Output directory should be empty."
        }
      } finally {
        if (Test-Path $in_dir) {
          Remove-Item -Recurse -Force $in_dir
        }
        if (Test-Path $out_dir) {
          Remove-Item -Recurse -Force $out_dir
        }
      }
    }
  },
  @{
    Name = "Different rename_format"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in"
      $out_dir = Join-Path $base_dir "tmp\out"
      $test_img_src = Join-Path $base_dir "tmp\test.jpg"
      $test_img = Join-Path $in_dir "test.jpg"
      $out_file = Join-Path $out_dir "2025/11/22/12-34-56.jpg"
      $tmp_cfg_path = Join-Path $base_dir "tmp\custom_cfg.json"
      $cfg_content = @{
        recursive     = $false
        rename_format = "{yyyy}/{mm}/{dd}/{hh}-{mi}-{ss}.{ext}"
      } | ConvertTo-Json

      try {
        # Create a dummy jpg file
        $bitmap = New-Object System.Drawing.Bitmap(1, 1)
        $bitmap.SetPixel(0, 0, [System.Drawing.Color]::Red)
        $bitmap.Save($test_img_src, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Dispose()
        exiftool "-DateTimeOriginal=2025:11:22 12:34:56" -overwrite_original $test_img_src | Out-Null

        New-Item -ItemType Directory -Force -Path $in_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $out_dir | Out-Null
        Copy-Item -Path $test_img_src -Destination $test_img
        Set-Content -Path $tmp_cfg_path -Value $cfg_content

        & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $tmp_cfg_path
        if (!(Test-Path $out_file)) {
          throw "Test failed: Output file not found."
        }
      } finally {
        if (Test-Path $in_dir) {
          Remove-Item -Recurse -Force $in_dir
        }
        if (Test-Path $out_dir) {
          Remove-Item -Recurse -Force $out_dir
        }
        if (Test-Path $test_img_src) {
          Remove-Item -Force $test_img_src
        }
        if (Test-Path $tmp_cfg_path) {
          Remove-Item -Force $tmp_cfg_path
        }
      }
    }
  },
  @{
    Name = "Recursive option"
    Test = {
      $in_dir = Join-Path $base_dir "tmp\in"
      $out_dir = Join-Path $base_dir "tmp\out"
      $test_img_src = Join-Path $base_dir "tmp\test.jpg"
      $sub_dir = Join-Path $in_dir "sub"
      $test_img1 = Join-Path $in_dir "test1.jpg"
      $test_img2 = Join-Path $sub_dir "test2.jpg"
      $out_file1 = Join-Path $out_dir "20251122/20251122_test1.jpg"
      $out_file2 = Join-Path $out_dir "20251122/20251122_test2.jpg"
      $tmp_cfg_path = Join-Path $base_dir "tmp\recursive_cfg.json"
      $cfg_content = @{
        recursive     = $true
        rename_format = "{yyyymmdd}/{yyyymmdd}_{filename}.{ext}"
      } | ConvertTo-Json

      try {
        # Create a dummy jpg file
        $bitmap = New-Object System.Drawing.Bitmap(1, 1)
        $bitmap.SetPixel(0, 0, [System.Drawing.Color]::Red)
        $bitmap.Save($test_img_src, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Dispose()
        exiftool "-DateTimeOriginal=2025:11:22 12:34:56" -overwrite_original $test_img_src | Out-Null

        New-Item -ItemType Directory -Force -Path $in_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $out_dir | Out-Null
        New-Item -ItemType Directory -Force -Path $sub_dir | Out-Null
        Copy-Item -Path $test_img_src -Destination $test_img1
        Copy-Item -Path $test_img_src -Destination $test_img2
        Set-Content -Path $tmp_cfg_path -Value $cfg_content

        & $script_path -in_dir $in_dir -out_dir $out_dir -cfg $tmp_cfg_path
        if (!(Test-Path $out_file1)) {
          throw "Test failed: Output file 1 not found."
        }
        if (!(Test-Path $out_file2)) {
          throw "Test failed: Output file 2 not found."
        }
      } finally {
        if (Test-Path $in_dir) {
          Remove-Item -Recurse -Force $in_dir
        }
        if (Test-Path $out_dir) {
          Remove-Item -Recurse -Force $out_dir
        }
        if (Test-Path $test_img_src) {
          Remove-Item -Force $test_img_src
        }
        if (Test-Path $tmp_cfg_path) {
          Remove-Item -Force $tmp_cfg_path
        }
      }
    }
  }
)

$passed = 0
$failed = 0

foreach ($test in $tests) {
  Write-Host "Running test: $($test.Name)"
  try {
    & $test.Test
    Write-Host "  [PASS]"
    $passed++
  } catch {
    Write-Host "  [FAIL] $_"
    $failed++
  }
}

Write-Host "--------------------"
Write-Host "Tests passed: $passed"
Write-Host "Tests failed: $failed"

if ($failed -gt 0) {
  exit 1
}
