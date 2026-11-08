# SmartMediaSorter & Find-MediaDuplicates

A PowerShell toolkit for automated photo sorting, renaming based on **EXIF** capture date (`SmartMediaSorter.ps1`), and media duplicate detection & isolation (`Find-MediaDuplicates.ps1`), backed by a shared hashing module and persistent JSON database (`SmartMediaCommon.ps1` / `.smartmediasorter_db.json`).

> Polska wersja dokumentacji: [README_PL.md](README_PL.md) (lub [REDME_PL.md](REDME_PL.md)).

---

## Requirements

- **PowerShell 7+** (`pwsh`). Windows PowerShell 5.1 (`powershell.exe`) is **not supported**; running scripts under 5.1 will terminate with a parse error.
- Optional: `ExifTool` in `PATH`, needed only for `SmartMediaSorter.ps1`. When absent, the script falls back to its built-in .NET EXIF reader.

Run scripts in `pwsh`, e.g.:
```powershell
pwsh -File .\Remove-DuplicateFolders.ps1 -WhatIf
```

---

## Architecture Overview

- **`SmartMediaCommon.ps1`** – Shared module providing file `SHA-256` hashing (`Get-CachedFileHash`), directory state hashing (`Get-DirectoryStateHash`), base name normalization (`Get-NormalizedBaseName`), logging (`Write-SorterLog`), and JSON database persistence (`.smartmediasorter_db.json`).
- **`SmartMediaSorter.ps1`** – Primary script for sorting and moving (or copying) photos from a source directory (`X:\Photos`) into a structured destination hierarchy (`X:\Pictures\YYYY\YYYY-MM\YYYY-MM-DD_HHMMSS_XXX.<ext>`).
- **`Find-MediaDuplicates.ps1`** – Script for duplicate detection (first by normalized base name and identical extension, then by byte length and `SHA-256` checksum) and moving confirmed duplicates into a local `_Duplicates` subfolder in the same directory.
- **`Remove-DuplicateFolders.ps1`** – Script for recursively removing `_Duplicates` folders to the Windows Recycle Bin (local drives) or to a dated `_DuplicatesTrash` directory (network/removable drives).

### Automatic Recycle Bin & System Folder Exclusion
All scripts (`SmartMediaSorter.ps1`, `Find-MediaDuplicates.ps1`, `Remove-DuplicateFolders.ps1`) leverage `SmartMediaCommon.ps1` to automatically ignore:
- **Windows Recycle Bin:** `$RECYCLE.BIN`, `RECYCLER`, `RECYCLED`
- **Network / NAS Recycle Bins:** `#recycle` (Synology DSM), `@Recycle` (QNAP QTS)
- **Unix / macOS Trash:** `.Trash`, `.Trash-*`, `.Trashes`
- **Project Trash Directory:** `_DuplicatesTrash`
- **System Folders:** `System Volume Information`

These folders are skipped during directory tree traversal without descending into them, preventing the processing of deleted files and eliminating Windows permission errors. Passing any such folder directly as an input parameter (`-Path`, `-SourcePath`, `-DestinationPath`) results in a safe and controlled refusal to run.

---

## 1. Photo Sorting Script (`SmartMediaSorter.ps1`)

### Key Features
1. **Target Structure & File Naming:**
   - Source directory (default): `X:\Photos`
   - Destination directory (default): `X:\Pictures`
   - Subfolder structure: `\YYYY\YYYY-MM\`
   - File name format: `YYYY-MM-DD_HHMMSS_XXX.<original_extension>` (lowercase)
   - Example: `X:\Pictures\2015\2015-08\2015-08-30_114023_000.jpg`
2. **Date Extraction (3-tier mechanism):**
   - **Step 1 (ExifTool):** Batch extraction of `DateTimeOriginal`, `CreateDate`, `MediaCreateDate`, `TrackCreateDate` tags via `ExifTool` (supporting JPEG, HEIC, RAW, TIFF, PNG, etc.).
   - **Step 2 (Native .NET EXIF Reader):** If `ExifTool` is not installed, the script directly parses binary EXIF headers (`APP1` / `TIFF`: `DateTimeOriginal 0x9003`, `CreateDate 0x9004`, `ModifyDate 0x0132`).
   - **Step 3 (Fallback):** If a file contains no EXIF metadata, the script falls back to the file modification timestamp (`LastWriteTime`) and logs the occurrence.
3. **Duplicate Detection & Burst Photos (`_XXX`):**
   - If a file with the same timestamp second (`YYYY-MM-DD_HHMMSS_*.ext`) already exists in the destination directory, file size and **SHA-256** checksum are compared:
     - **Identical SHA-256 (Duplicate):** The file is skipped (left in source `X:\Photos`) and logged as `DUPLICATE`.
     - **Different SHA-256 (e.g. burst capture in the same second):** The file receives the next available three-digit sequence number (`_001`, `_002`, etc.).
4. **Directory State & Hash Database (`.smartmediasorter_db.json`):**
   - Creates and updates a database in the destination directory (`<DestinationPath>\.smartmediasorter_db.json`).
   - On subsequent runs, unchanged files and subdirectories reuse cached data (skipping redundant `ExifTool` calls and disk `SHA-256` recalculations).

### Usage Examples (`SmartMediaSorter.ps1`)
```powershell
# Dry run simulation (-WhatIf)
.\SmartMediaSorter.ps1 -WhatIf

# Default move from X:\Photos to X:\Pictures
.\SmartMediaSorter.ps1

# Copy instead of moving
.\SmartMediaSorter.ps1 -CopyOnly

# Force a full database rebuild
.\SmartMediaSorter.ps1 -RebuildDatabase
```

---

## 2. Duplicate Detection Script (`Find-MediaDuplicates.ps1`)

### How It Works
1. **Different Extensions = Different Files:** Files with different extensions (e.g., `.jpg` vs `.png` or `.cr2` vs `.jpg`) are never considered duplicates.
2. **Step 1 – Name-Based Hinting:** Files sharing the same extension are grouped by normalized base name (`Get-NormalizedBaseName`), stripping common copy suffixes and sequence numbers (e.g., `DSC_5000.jpg`, `DSC_5000 (1).jpg`, `DSC_5000 - Copy.jpg`, `DSC_5000_copy.jpg`, as well as `2015-08-30_114023_000.jpg` and `2015-08-30_114023_001.jpg`).
3. **Step 2 – Content Verification (Size + SHA-256):** Within each name group, byte sizes (`Length`) are compared, followed by computing or querying the `SHA-256` checksum from `.smartmediasorter_db.json`.
4. **Step 3 – Move to Local `_Duplicates` Folder:** The file with the cleanest/shortest base name (e.g., `DSC_5000.jpg`) is preserved in place as the original, while all confirmed duplicates are moved to a local `_Duplicates` subfolder created directly in the directory where each duplicate was found (e.g., `X:\Photos\AlbumA\_Duplicates\`).

### Usage Examples (`Find-MediaDuplicates.ps1`)
```powershell
# Dry run simulation (-WhatIf) in default directory X:\Photos
.\Find-MediaDuplicates.ps1 -WhatIf

# Detect and move duplicates in X:\Photos to local _Duplicates folders
.\Find-MediaDuplicates.ps1 -Path "X:\Photos"

# Scan X:\Pictures with a custom duplicate folder name
.\Find-MediaDuplicates.ps1 -Path "X:\Pictures" -DuplicateFolderName "_Duplikaty"
```

### Parameters (`Find-MediaDuplicates.ps1`)

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `-Path` | `string` | `"X:\Photos"` | Directory to scan for duplicate files. |
| `-DuplicateFolderName` | `string` | `"_Duplicates"` | Subfolder name created in duplicate's directory where duplicates are moved. |
| `-Recurse` | `bool` | `$true` | Recursively scan subdirectories (automatically skipping `_Duplicates` folders). |
| `-DatabasePath` | `string` | `<Path>\.smartmediasorter_db.json` | Path to shared JSON hash database. |
| `-NoDatabase` | `switch` | `$false` | Disables database read and write. |
| `-RebuildDatabase` | `switch` | `$false` | Ignores cached database hashes and recalculates from disk. |
| `-IncludeExtensions` | `string[]` | `@()` (all files) | Optional list of file extensions to filter by. |
| `-ThrottleLimit` | `int` | `8` | Maximum concurrent threads (`RunspacePool`) scanning subdirectories and computing hashes (`1` = single-threaded). |
| `-DirectoryTimeoutSeconds` | `int` | `300` | Maximum processing time per subdirectory in seconds; stuck threads are terminated and logged with the current file name (`0` = no timeout). |
| `-LogPath` | `string` | `FindMediaDuplicates_<timestamp>.log` | Path to log file. |
| `-WhatIf` | `switch` | `$false` | Simulation mode – displays detected duplicates without moving files. |

---

## 3. Duplicate Folders Cleanup Script (`Remove-DuplicateFolders.ps1`)

Recursively discovers all `_Duplicates` folders (created by `Find-MediaDuplicates.ps1`) within a given path and removes them along with their contents in a **recoverable manner**.

### How It Works
1. **Discovery:** Traverses the directory tree without following reparse points (junctions/symlinks). Folder names are matched case-insensitively, and the script does not descend into found `_Duplicates` folders.
2. **Removal Strategy (`-Mode Auto`, default):**
   - **Local Drive (`Fixed`)** – Folders are moved to the **Windows Recycle Bin**.
   - **Network Drive (e.g. `X:`), Removable, or UNC Path** – Windows lacks Recycle Bin support on non-fixed drives (files would be permanently deleted), so folders are safely **moved** (fast volume-local rename) to:
     `<root>\_DuplicatesTrash\<yyyyMMdd_HHmmss>\<relative_path>`, e.g. `X:\_DuplicatesTrash\20261108_142152\Photos\2020\Album\_Duplicates`.
     After verifying the contents, the `X:\_DuplicatesTrash` folder can be deleted manually.
3. **Safety Measures:**
   - Prompts for confirmation before each folder (`ConfirmImpact = High`; answering `A` confirms "Yes to All"). The `-Force` switch suppresses prompts.
   - `-WhatIf` displays the list of candidate folders with file count and total size without deleting anything.
   - The trash directory is excluded from scanning, and `_Duplicates` folders that are symbolic links or junctions are skipped with a warning.
   - Refuses to execute if `-Path` itself is a `_Duplicates` folder.
   - Deleting duplicate folders requires no changes to `.smartmediasorter_db.json`, as `Find-MediaDuplicates.ps1` removes file records immediately upon moving them to `_Duplicates`.

### Usage Examples (`Remove-DuplicateFolders.ps1`)
```powershell
# Preview folders to be removed in X:\Photos (default)
.\Remove-DuplicateFolders.ps1 -WhatIf

# Remove with interactive confirmation for each folder
.\Remove-DuplicateFolders.ps1

# Remove without prompts in another directory
.\Remove-DuplicateFolders.ps1 -Path "X:\Pictures" -Force
```

### Parameters (`Remove-DuplicateFolders.ps1`)

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `-Path` | `string` | `"X:\Photos"` | Directory to search recursively. |
| `-DuplicateFolderName` | `string` | `"_Duplicates"` | Name of folders to remove. |
| `-Mode` | `Auto` / `RecycleBin` / `TrashFolder` | `Auto` | Deletion mode (`RecycleBin` is only permitted on local drives). |
| `-TrashPath` | `string` | `<root>\_DuplicatesTrash` | Trash directory path for `TrashFolder` mode (must reside on the same volume). |
| `-Force` | `switch` | `$false` | Suppresses confirmation prompts. |
| `-ThrottleLimit` | `int` | `8` | Maximum concurrent threads scanning top-level subdirectories and sizing `_Duplicates` folders. |
| `-LogPath` | `string` | `RemoveDuplicateFolders_<timestamp>.log` | Path to log file. |
| `-PassThru` | `switch` | `$false` | Returns an object containing execution statistics. |
| `-WhatIf` | `switch` | `$false` | Simulation mode. |

---

## Running Automated Tests

The integration test suite verifies all tools (`SmartMediaSorter.ps1`, `Find-MediaDuplicates.ps1`, `Remove-DuplicateFolders.ps1`) alongside the shared module `SmartMediaCommon.ps1`:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-SmartMediaSorter.ps1
```
