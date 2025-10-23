# Structured Paste

![Icon](https://img.shields.io/badge/OS-Windows-blue) ![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-orange) ![License](https://img.shields.io/badge/License-MIT-green)

A Windows utility that recreates the directory structure of copied files or folders when pasting, starting from the earliest common parent folder. This avoids the need to manually recreate folder hierarchies in the destination.

## Quick Start: Why Use This?

If you're tired of recreating folder structures manually when moving files around, this script provides a context menu option ("Paste with Structure") to handle it automatically. It works for single files, multiple files, or entire folders (including subfolders and their contents).

## Features

- Recreates the source directory tree in the destination based on the common prefix between source and destination paths. (Note: For pastes across different drives, it recreates the full relative path from the source drive root, excluding the drive letter.)
- Provides conflict resolution for existing files via a dialog with options: Overwrite, Rename (auto-appends (1), (2), etc.), Skip, or Abort, plus an "Apply to all" checkbox for multiple conflicts.
- Supports Unicode and special characters in paths.
- No external dependencies; uses built-in PowerShell and .NET assemblies.

## Installation

This is a Windows-only script (tested on Windows 10+). Place the files anywhere convenient and add a registry entry for the context menu. Registering the context menu entry requires admin privileges for system-wide changes. If you're not an admin, see the per-user option in Advanced Configuration.

**Where to Place Scripts?**
- **Anywhere is fine**: Use a custom folder like `C:\Users\YourName\Tools` if you prefer.
- **Why/Why Not C:\Windows?**
  - **Pros**: Files are less likely to be accidentally moved or deleted, ensuring the registry links remain valid. Additionally, if placed there, you can run the script from the command line for non-context-menu use, e.g., `structured_paste .` (current directory) or `structured_paste folder1` (specific folder).
  - **Cons**: Requires admin privileges to write there; clutters the system folder.

### Step 1: Download the Scripts
Save these files in a folder of your choice:
- `structured_paste.ps1`: The main PowerShell script.
- `structured_paste.vbs`: A VBS wrapper to run PowerShell without showing a console window.
- `structured_paste_register.reg`: Registry file to add the context menu.
- `structured_paste_unregister.reg`: To remove it later.

### Step 2: Edit Paths (If Needed)
- In `structured_paste_register.reg`: Replace `C:\\Windows` with your folder path, using double backslashes for directory separators (e.g., `C:\\Users\\YourName\\Tools`). Note: Registry files require this escaping for paths in string values. Avoid invalid characters in paths (e.g., no unescaped / in key names). For more on valid registry syntax and escaping, see [Microsoft's Registry Guide](https://learn.microsoft.com/en-us/windows/win32/sysinfo/structure-of-the-registry).
- In `structured_paste.vbs`: Update the PowerShell file path if not in `C:\Windows`.

### Step 3: Register the Context Menu
1. Double-click `structured_paste_register.reg`.
2. If prompted, allow admin access (UAC prompt).
3. Confirm the addition in the dialog.

The "Paste with Structure" option should now appear when you right-click in a folder's empty space. On Windows 11, if it's not in the default menu, use Shift + right-click or click "Show more options" to access the classic menu.

To remove: Double-click `structured_paste_unregister.reg`.

## Usage

1. Copy files or folders as usual (Ctrl+C or right-click > Copy). Supports single items, multiple files, or folders (with subfolders and their contents).
2. Navigate to your destination folder.
3. Right-click in the empty area (folder background) > "Paste with Structure." (On Windows 11, if not visible, look under "Show more options")
4. If conflicts occur (e.g., file exists), a dialog appears with the source/destination paths and options.
5. Folders are created as needed, and files are copied with the structure preserved.

If nothing happens: First, ensure the clipboard contains copied files or folders (not text or other data). Then, check the log at `%TEMP%\structuredpaste.log` for details. If issues persist, see Troubleshooting.

### Examples:
Pull only essential reports from a labyrinthine company drive for an audit; pasting reconstructs the department/project hierarchy in your secure archive in seconds.
- Source: `S:\Company\Finance\Projects\Q3_2025\budget_report.xlsx`
- Destination: `F:\Archives\Audit2025\`
- Result: `F:\Archives\Audit2025\Company\Finance\Projects\Q3_2025\budget_report.xlsx`

Select multiple edited photos and folders from across your sprawling photo archive; pasting them into your external drive rebuilds the exact year/event/location folders instantly.
- Source: `C:\Photos\2023\Vacation\Italy\Rome\closeups` + `group_shot.jpg` + `colosseum_edit.jpg,` + ... 
- Destination: `E:\PhotosForPrinting\`
- Result: `E:\PhotosForPrinting\Photos\2023\Vacation\Italy\Rome\closeups` + `group_shot.jpg` + `colosseum_edit.jpg` + ...

## Advanced Configuration

- **Customize Menu Text/Icon**: Edit `@="Paste with Structure"` and `"Icon"="powershell.exe"` in the `.reg` file.
- **Per-User Installation (No Admin Privileges Needed)**: Change the registry key in `structured_paste_register.reg` to `HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell\StructuredPaste` instead of `HKEY_CLASSES_ROOT`. This adds the menu only for your user account.
- **Menu Placement**: By default, the entry appears in the primary context menu. To make it appear only in the extended menu (Shift + right-click), add `"Extended"=""` under the `[HKEY_CLASSES_ROOT\Directory\Background\shell\StructuredPaste]` section in the `structured_paste_register.reg` file.
- **Command-Line Use**: If the script is in `C:\Windows` (or your PATH), run it directly: `structured_paste .` for the current directory or `structured_paste "C:\Path\To\Folder"` for a specific destination.

## Risks and Warnings

This script reads from sources but writes to destinations, similar to standard copy-paste operations.

**Potential Data Loss/Damage:**
- Choosing "Overwrite" (especially with "Apply to all") replaces existing files without recovery. No built-in undo. Test on non-critical data first; always backup important data.
- Bugs could result in incomplete copies or empty folders, but no intentional deletion occurs outside chosen overwrites.

Use at your own risk; not liable for data loss/damage.

## Troubleshooting and Issue Reporting

Common issues:
- Menu missing: Re-run `structured_paste_register.reg` as admin or check for typos in paths.
- "Invalid path" error: Ensure destination is a valid folder.
- No action on paste: Clipboard may be empty or contain non-file data.

The script includes logging to `%TEMP%\structuredpaste.log` for debug info, such as paths processed and errors.

**Edge Cases Handled:**
- Path too long: Warns and skips (files >259 chars, folders >247 to allow for subfiles).
- Pasting into source/subtree: Detects and aborts with a message listing affected items.
- Empty clipboard: Shows an info message and exits.

**Reporting Issues:**
Open an issue on this repo with details. Use this template:

```
### Issue Description
[Briefly describe the problem, e.g., "Script fails on cross-drive paste with error X."]

### Steps to Reproduce
1. [Step 1]
2. [Step 2]
...

### Expected Behavior
[What should happen?]

### Actual Behavior
[What happened instead?]

### Environment
- Windows Version: [e.g., Windows 11 22H2]
- PowerShell Version: [Run `$PSVersionTable.PSVersion`]
- Script Location: [e.g., C:\Tools]
- Admin/Non-Admin: [Yes/No]

### Log File
[Paste relevant lines from %TEMP%\structuredpaste.log, or attach the file.]

### Screenshots (if applicable)
[Add images of errors/dialogs.]

### Additional Context
[Any other info, e.g., special characters in paths?]
```

## Final Notes
- **Security**: Scripts are plain text—review for safety.
- **Updates**: Check the repo for improvements. Test after OS updates.
- **License**: MIT—free to use, modify, and share.
- **Credits**: Built with PowerShell; inspired by common file management pains.


If searching for tools like this, terms such as — paste with subtree, copy with directory structure, paste with folders, paste with folder chain, clone directory on paste, pre-create directory on copy, structured paste utility, preserve folder hierarchy on paste, recursive copy with relative paths, windows structured copy-paste, folder tree paste, subdirectory paste tool, path-preserving paste, hierarchy-aware copy, deep copy paste, nested folder paste, tree structure paste, directory recreation on paste, subfolder paste script, folder path clone, structured file transfer, paste retaining folders, windows explorer structured paste, folder structure duplicator, path tree copier, hierarchical paste addon, subtree copy utility, or directory chain paste — might help.
