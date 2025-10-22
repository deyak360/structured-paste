# self-paste-with-structure.ps1

<#
Structured Paste Utility

Description:
This PowerShell script enhances the standard Windows copy-paste functionality by recreating the directory structure of copied files or folders relative to their earliest common parent when pasting. Instead of manually recreating files in the destination folder, it preserves the subtree hierarchy starting from the point where the source and destination paths diverge.

What it Achieves:
- Allows structured pasting via a context menu option ("Paste with Structure") when right-clicking in a folder's background.
- Handles single files, multiple files, or entire folders (including subfolders and files).
- Computes the relative path based on the common prefix between source and destination.
- If source and destination are on different drives, it recreates the full path from the source drive root (excluding the drive letter).
- Supports Unicode and special characters in paths.
- Checks for MAX_PATH (260 characters) limitations and warns if exceeded.
- Prevents pasting into the source or its subtree to avoid infinite loops or overwrites.
- Handles permission errors by displaying warnings.
- Provides conflict resolution for existing files with options: Skip, Rename (auto-appends (1), (2), etc.), Overwrite, Cancel All.
- Includes an "Apply to all conflicts" checkbox in the dialog for batch operations.

Dependencies/Assumptions:
- Requires PowerShell (available on Windows by default).
- Uses .NET assemblies: System.Windows.Forms (for dialogs and clipboard access) and Microsoft.VisualBasic (though not explicitly used here, loaded for potential future use).
- Integrated via registry for context menu: Uses a VBS wrapper to run PowerShell hidden (no console popup).
- Clipboard must contain file paths (via Ctrl+C or right-click Copy on files/folders).
- Works on local drives; network paths may work but not tested extensively.
- Assumes Windows environment; paths use backslashes.
- Logging to %TEMP%\structuredpaste.log for debugging.
- No external dependencies; all operations use built-in cmdlets.

How Files/Folders are Handled:
- Clipboard items are retrieved as a list of paths.
- For each item:
  - If file: Computes relative path from parent directory.
  - If folder: Computes relative path from the folder's parent, treating the copied folder as the leaf.
- Creates necessary directories in the destination.
- Copies files/folders recursively if a folder.
- Cross-drive: Recreates full relative path from source root.
- Multiple items: Processed sequentially; conflicts handled per item or "apply to all".

Edge-Cases:
- Drive roots as destination (e.g., W:\): Handled by appending '\' if needed; registry uses "%V\." to avoid quote escaping issues in roots like "W:\".
- Empty clipboard: Shows info message and exits.
- Invalid destination: Errors out with message.
- Path too long: Warns and skips (files >259 chars, folders >247 to allow for files inside).
- Pasting into source/subtree: Detects and aborts with list of conflicting items.
- Permission errors: Caught in try-catch, shows error message.
- Unicode/special chars: Handled natively by PowerShell/.NET.
- If common prefix is empty (different drives), uses path from source root.

Conflict-Resolution:
- Before copying a file, checks if target exists.
- Shows a dialog with source/dest paths and options: Overwrite, Rename, Skip, Cancel All.
- "Apply to all" checkbox: If checked, applies choice to subsequent conflicts without prompting.
- Rename: Appends (1), (2), etc., to filename.
- Folders: Conflicts handled during recursive copy for files inside; folder creation is forced if needed.
- Note: Does not merge folders; overwrites files only if chosen.

Installation:
- Run the .reg file to add context menu.
- Place in C:\Windows

Reasons for Odd Choices:
- Registry uses "%V\." instead of "%V\": The dot handles drive roots without escaping quotes (e.g., "C:\" becomes "C:\." which is valid and equivalent).
- VBS wrapper: Hides PowerShell window; direct registry call shows console.
- Separate loop for subtree check: Collects all conflicts first to show a single message.
- Global $script:applyToAll: Persists choice across items for "apply to all".
- Throw "Canceled": Simple way to propagate abort up the call stack.

#>

param(
    [string]$destination
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$logFile = Join-Path $env:TEMP "structuredpaste.log"
Add-Content -Path $logFile -Value "----- New run $(Get-Date) -----"
Add-Content -Path $logFile -Value "Raw param: '$destination'"

$destination = $destination.Trim()  # Remove leading/trailing whitespace
Add-Content -Path $logFile -Value "After trim: '$destination'"

$destination = $destination.TrimEnd([System.IO.Path]::DirectorySeparatorChar)  # Remove trailing backslash if present
Add-Content -Path $logFile -Value "After trimEnd: '$destination'"

if ($destination -match '^[a-zA-Z]:$') {
    $destination += [System.IO.Path]::DirectorySeparatorChar  # Add backslash for drive roots (e.g., C: -> C:\)
    Add-Content -Path $logFile -Value "After match add: '$destination'"
}

Add-Content -Path $logFile -Value "Path for GetFullPath: '$destination'"

try {
    $destination = [System.IO.Path]::GetFullPath($destination)  # Normalize to absolute path
    Add-Content -Path $logFile -Value "Full destination: '$destination'"
} catch {
    Add-Content -Path $logFile -Value "Error in GetFullPath: $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show("Invalid destination path: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

if (!(Test-Path $destination -PathType Container)) {
    Add-Content -Path $logFile -Value "Destination not container"
    [System.Windows.Forms.MessageBox]::Show("Destination path does not exist or is not a directory.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

$clipboardItems = [System.Windows.Forms.Clipboard]::GetFileDropList()  # Get copied file/folder paths
if ($clipboardItems.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No files or folders in the clipboard.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    exit
}

$script:applyToAll = $null

function Get-CommonPrefix {
    param(
        [string]$path1,
        [string]$path2
    )

    # Get drive roots
    $root1 = [System.IO.Path]::GetPathRoot($path1)
    $root2 = [System.IO.Path]::GetPathRoot($path2)
    if ($root1 -ne $root2) {
        return ""  # Different drives: no common prefix
    }

    # Normalize paths
    $path1 = [System.IO.Path]::GetFullPath($path1).TrimEnd([System.IO.Path]::DirectorySeparatorChar)
    $path2 = [System.IO.Path]::GetFullPath($path2).TrimEnd([System.IO.Path]::DirectorySeparatorChar)

    # Split into parts
    $parts1 = $path1 -split '\\'
    $parts2 = $path2 -split '\\'

    # Find matching prefix length
    $minLen = [Math]::Min($parts1.Length, $parts2.Length)
    $i = 0
    while ($i -lt $minLen -and $parts1[$i] -ieq $parts2[$i]) {
        $i++
    }

    if ($i -eq 0) {
        return ""  # No common parts
    }

    return $parts1[0..($i-1)] -join '\'  # Join matching parts
}

function Get-RelativePath {
    param(
        [string]$sourceDir,
        [string]$commonPrefix
    )

    if ([string]::IsNullOrEmpty($commonPrefix)) {
        $rootLen = [System.IO.Path]::GetPathRoot($sourceDir).Length
        return $sourceDir.Substring($rootLen)  # Full relative from root if no common
    } else {
        return $sourceDir.Substring($commonPrefix.Length).TrimStart('\')  # Relative after prefix
    }
}

function Show-ConflictDialog {
    param(
        [string]$sourcePath,
        [string]$targetPath
    )

    # Create form for conflict resolution
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "File Exists"
    $form.Size = New-Object System.Drawing.Size(465, 220)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    # Label with details
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "The file '$targetPath' already exists.`n`nSource: $sourcePath`nDestination: $targetPath"
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(420, 60)
    $form.Controls.Add($label)

    # Buttons
    $btnOverwrite = New-Object System.Windows.Forms.Button
    $btnOverwrite.Text = "Overwrite"
    $btnOverwrite.Location = New-Object System.Drawing.Point(10, 80)
    $btnOverwrite.Size = New-Object System.Drawing.Size(100, 30)
    $btnOverwrite.Add_Click({ $form.Tag = @{Choice="Overwrite"; ApplyAll=$checkbox.Checked}; $form.Close() })
    $form.Controls.Add($btnOverwrite)

    $btnRename = New-Object System.Windows.Forms.Button
    $btnRename.Text = "Rename"
    $btnRename.Location = New-Object System.Drawing.Point(120, 80)
    $btnRename.Size = New-Object System.Drawing.Size(100, 30)
    $btnRename.Add_Click({ $form.Tag = @{Choice="Rename"; ApplyAll=$checkbox.Checked}; $form.Close() })
    $form.Controls.Add($btnRename)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text = "Skip"
    $btnSkip.Location = New-Object System.Drawing.Point(230, 80)
    $btnSkip.Size = New-Object System.Drawing.Size(100, 30)
    $btnSkip.Add_Click({ $form.Tag = @{Choice="Skip"; ApplyAll=$checkbox.Checked}; $form.Close() })
    $form.Controls.Add($btnSkip)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abort"
    $btnCancel.Location = New-Object System.Drawing.Point(340, 80)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Add_Click({ $form.Tag = @{Choice="Cancel"; ApplyAll=$false}; $form.Close() })
    $form.Controls.Add($btnCancel)

    # Checkbox for apply to all
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Text = "Apply to all conflicts"
    $checkbox.Location = New-Object System.Drawing.Point(10, 120)
    $checkbox.Size = New-Object System.Drawing.Size(200, 30)
    $form.Controls.Add($checkbox)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Get-UniqueName {
    param(
        [string]$path  # Existing path to rename
    )

    if (!(Test-Path $path)) {
        return $path  # No conflict
    }

    $dir = Split-Path $path -Parent
    $base = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $ext = [System.IO.Path]::GetExtension($path)
    $counter = 1

    do {
        $newPath = Join-Path $dir "$base ($counter)$ext"
        $counter++
    } while (Test-Path $newPath)

    return $newPath
}

function RecursiveCopy {
    param(
        [string]$source,
        [string]$target
    )

    try {
        if (!(Test-Path $target)) {
            New-Item -Path $target -ItemType Directory -Force | Out-Null  # Create target dir
        }

        # Recurse through items
        Get-ChildItem -Path $source | ForEach-Object {
            $srcItem = $_.FullName
            $tgtItem = Join-Path $target $_.Name

            if ($_.PSIsContainer) {
                RecursiveCopy -source $srcItem -target $tgtItem  # Recurse for subfolders
            } else {
                if ($tgtItem.Length -gt 259) {
                    [System.Windows.Forms.MessageBox]::Show("Path too long for file '$srcItem': $tgtItem", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return  # Skip long paths
                }

                if (Test-Path $tgtItem) {
                    if ($script:applyToAll) {
                        $choice = $script:applyToAll
                    } else {
                        $result = Show-ConflictDialog -sourcePath $srcItem -targetPath $tgtItem
                        $choice = $result.Choice
                        if ($result.ApplyAll -and $choice -ne "Cancel") {
                            $script:applyToAll = $choice
                        }
                    }
                    switch ($choice) {
                        "Skip" { return }
                        "Overwrite" { Copy-Item -Path $srcItem -Destination $tgtItem -Force }
                        "Rename" {
                            $newTgt = Get-UniqueName -path $tgtItem
                            Copy-Item -Path $srcItem -Destination $newTgt -Force
                        }
                        "Cancel" { throw "Canceled" }
                    }
                } else {
                    Copy-Item -Path $srcItem -Destination $tgtItem -Force  # Copy file
                }
            }
        }
    } catch {
        if ($_.Exception.Message -eq "Canceled") {
            throw "Canceled"  # Propagate abort
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            throw  # Rethrow for outer catch
        }
    }
}

$conflictingItems = New-Object System.Collections.ArrayList

# First pass: Check for subtree pastes
foreach ($item in $clipboardItems) {
    if (!(Test-Path $item)) {
        continue
    }

    $isFile = (Test-Path $item -PathType Leaf)

    if ($isFile) {
        $sourceDir = Split-Path $item -Parent
        $leaf = Split-Path $item -Leaf
    } else {
        $sourceDir = $item.TrimEnd([System.IO.Path]::DirectorySeparatorChar)
        $leaf = Split-Path $sourceDir -Leaf
        $sourceDir = Split-Path $sourceDir -Parent
    }

    $commonPrefix = Get-CommonPrefix -path1 $sourceDir -path2 $destination
    $relativePath = Get-RelativePath -sourceDir $sourceDir -commonPrefix $commonPrefix

    $targetDir = Join-Path $destination $relativePath
    $targetPath = Join-Path $targetDir $leaf

    if ($isFile) {
        if ($targetPath -ieq $item) {
            [void]$conflictingItems.Add($item)  # Same file
        }
    } else {
        if ($targetPath -ieq $item -or $targetPath.StartsWith($item + [System.IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            [void]$conflictingItems.Add($item)  # Folder or subfolder
        }
    }
}

if ($conflictingItems.Count -gt 0) {
    $msg = "Cannot paste into the source location or subtree for the following items:`n`n" + ($conflictingItems -join "`n")
    [System.Windows.Forms.MessageBox]::Show($msg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Second pass: Perform copies
foreach ($item in $clipboardItems) {
    try {
        if (!(Test-Path $item)) {
            continue
        }

        $isFile = (Test-Path $item -PathType Leaf)

        if ($isFile) {
            $sourceDir = Split-Path $item -Parent
            $leaf = Split-Path $item -Leaf
        } else {
            $sourceDir = $item.TrimEnd([System.IO.Path]::DirectorySeparatorChar)
            $leaf = Split-Path $sourceDir -Leaf
            $sourceDir = Split-Path $sourceDir -Parent
        }

        Add-Content -Path $logFile -Value "SourceDir: '$sourceDir' Leaf: '$leaf'"

        $commonPrefix = Get-CommonPrefix -path1 $sourceDir -path2 $destination
        Add-Content -Path $logFile -Value "CommonPrefix: '$commonPrefix'"

        $relativePath = Get-RelativePath -sourceDir $sourceDir -commonPrefix $commonPrefix
        Add-Content -Path $logFile -Value "RelativePath: '$relativePath'"

        $targetDir = Join-Path $destination $relativePath
        $targetPath = Join-Path $targetDir $leaf
        Add-Content -Path $logFile -Value "TargetPath: '$targetPath'"

        if ($isFile) {
            if ($targetPath.Length -gt 259) {
                [System.Windows.Forms.MessageBox]::Show("Path too long for file '$item': $targetPath", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                continue
            }
        } else {
            if ($targetPath.Length -gt 247) {
                [System.Windows.Forms.MessageBox]::Show("Path too long for folder '$item': $targetPath", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                continue
            }
        }

        if ($isFile) {
            if (!(Test-Path $targetDir)) {
                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null  # Create dirs
            }

            if (Test-Path $targetPath) {
                if ($script:applyToAll) {
                    $choice = $script:applyToAll
                } else {
                    $result = Show-ConflictDialog -sourcePath $item -targetPath $targetPath
                    $choice = $result.Choice
                    if ($result.ApplyAll -and $choice -ne "Cancel") {
                        $script:applyToAll = $choice
                    }
                }
                switch ($choice) {
                    "Skip" { continue }
                    "Overwrite" { Copy-Item -Path $item -Destination $targetPath -Force }
                    "Rename" {
                        $newTarget = Get-UniqueName -path $targetPath
                        Copy-Item -Path $item -Destination $newTarget -Force
                    }
                    "Cancel" { throw "Canceled" }
                }
            } else {
                Copy-Item -Path $item -Destination $targetPath -Force
            }
        } else {
            RecursiveCopy -source $item -target $targetPath  # Handle folder recursively
        }
    } catch {
        Add-Content -Path $logFile -Value "Error processing '$item': $($_.Exception.Message)"
        if ($_.Exception.Message -eq "Canceled") {
            exit
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error processing '$item': $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}