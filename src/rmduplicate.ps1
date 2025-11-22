# No param() block. We will parse arguments manually from the $args array.

function Process-Paths {
  param(
    [string[]]$Paths
  )

  $files = Get-ChildItem -Path $Paths -Recurse -File

  $groups = $files | Group-Object -Property { (Get-FileHash $_.FullName -Algorithm MD5).Hash }

  foreach ($group in $groups) {
    if ($group.Count -gt 1) {
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

# --- Manual Argument Parsing ---
$userPaths = @()
$foundPathSwitch = $false
foreach ($arg in $args) {
    if ($foundPathSwitch) {
        $userPaths += $arg
    }
    if ($arg -eq '-Path') {
        $foundPathSwitch = $true
    }
}

if ($userPaths.Count -eq 0) {
    throw "Manual parser: Could not find any paths after the -Path switch. Arguments received: $($args -join ' ')"
}

Process-Paths -Paths $userPaths
