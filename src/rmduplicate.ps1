# No param() block. We will parse arguments manually from the $args array.

function Process-Paths {
  param(
    [string[]]$Paths,
    [switch]$Bulk
  )

  $files = Get-ChildItem -Path $Paths -Recurse -File
  $filesToDeleteInBulk = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

  $groups = $files | Group-Object -Property { (Get-FileHash $_.FullName -Algorithm MD5).Hash }

  foreach ($group in $groups) {
    if ($group.Count -gt 1) {
      if ($Bulk) {
        # --- Bulk Mode Logic ---
        Write-Host "Duplicate files found for hash $($group.Name):"
        $group.Group | ForEach-Object { Write-Host "  - $($_.FullName)" }

        # Sort by name length (shortest first), then by creation time (oldest first) as a tie-breaker.
        $sortedFiles = $group.Group | Sort-Object { $_.Name.Length }, CreationTime

        # The first file in the sorted list is the one to keep.
        $fileToKeep = $sortedFiles[0]
        Write-Host "  -> Keeping: $($fileToKeep.FullName)"

        # Add the rest to the deletion list.
        $filesToDeleteThisGroup = $sortedFiles | Select-Object -Skip 1
        foreach($file in $filesToDeleteThisGroup) {
          $filesToDeleteInBulk.Add($file)
          Write-Host "  -> Marked for deletion: $($file.FullName)"
        }
        Write-Host "" # Newline for separation
      } else {
        # --- Interactive Mode Logic ---
        Write-Host "Duplicate files found:"
        $group.Group | ForEach-Object { Write-Host "  - $($_.FullName)" }

        $choices = @()
        for ($i = 0; $i -lt $group.Group.Count; $i++) {
          $choices += New-Object System.Management.Automation.Host.ChoiceDescription -ArgumentList ("&" + ($i + 1) + " " + $group.Group[$i].FullName, "Delete " + $group.Group[$i].FullName)
        }
        $choices += New-Object System.Management.Automation.Host.ChoiceDescription -ArgumentList ("&" + ($choices.Count + 1) + " Skip", "Skip this set of duplicates")

        $title = "Select a file to delete"
        $message = "Which file would you like to delete?"
        $default = $choices.Count - 1

        $result = $Host.UI.PromptForChoice($title, $message, $choices, $default)

        if ($result -lt ($choices.Count - 1)) {
          $fileToDelete = $group.Group[$result]
          Write-Host "Moving $($fileToDelete.FullName) to Recycle Bin..."
          Add-Type -AssemblyName Microsoft.VisualBasic
          [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($fileToDelete.FullName, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
          Write-Host "File moved to Recycle Bin."
        }
      }
    }
  }

  # --- After the loop, handle bulk deletion confirmation ---
  if ($Bulk -and $filesToDeleteInBulk.Count -gt 0) {
    Write-Host "----------------------------------------"
    Write-Host "The following $($filesToDeleteInBulk.Count) files are marked for deletion:"
    $filesToDeleteInBulk | ForEach-Object { Write-Host "  - $($_.FullName)" }
    Write-Host "----------------------------------------"

    $choice = $Host.UI.PromptForChoice(
      "Confirm Deletion",
      "Do you want to move all these files to the Recycle Bin?",
      @("&Yes", "&No"),
      1 # Default to No
    )

    if ($choice -eq 0) {
      # User selected Yes
      Write-Host "Moving files to Recycle Bin..."
      Add-Type -AssemblyName Microsoft.VisualBasic
      foreach ($file in $filesToDeleteInBulk) {
        Write-Host "  -> Deleting $($file.FullName)"
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($file.FullName, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
      }
      Write-Host "Operation complete."
    } else {
      Write-Host "Deletion cancelled by user."
    }
  } elseif ($Bulk) {
    Write-Host "Bulk mode finished. No files marked for deletion."
  }
}

# --- Manual Argument Parsing ---
$userPaths = @()
$bulkSwitch = $false
$i = 0
while ($i -lt $args.Count) {
  switch -regex ($args[$i]) {
    '^-Path$' {
      $i++
      while ($i -lt $args.Count -and $args[$i] -notmatch '^-') {
        $userPaths += $args[$i]
        $i++
      }
      # The loop exits when it hits a switch or the end.
      # We need to decrement $i so the outer loop can process the switch.
      $i--
    }
    '^-Bulk$' {
      $bulkSwitch = $true
    }
    default {
      if ($args[$i] -notmatch '^-') {
        # Argument is not a switch. Could be a path.
        # To avoid confusion, we strictly require -Path.
        # So we just ignore it, and the check later will fail if -Path was not provided.
      } else {
        Write-Warning "Unknown argument: $($args[$i])"
      }
    }
  }
  $i++
}

if ($userPaths.Count -eq 0) {
  throw "Manual parser: The -Path parameter is required and must be followed by one or more paths. Example: -Path C:\folder1 C:\folder2"
}

Process-Paths -Paths $userPaths -Bulk:$bulkSwitch
