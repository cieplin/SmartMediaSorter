#Requires -Version 7.0

<#
.SYNOPSIS
    Recursively finds all duplicate quarantine folders (default: "_Duplicates") created by
    Find-MediaDuplicates.ps1 below a given directory and removes them together with their content
    in a recoverable way (Windows Recycle Bin or a dated trash folder on the same volume).

.DESCRIPTION
    The script walks the directory tree below -Path (without following junctions/symbolic links)
    and collects every folder whose name equals -DuplicateFolderName (case-insensitive). It does
    not descend into matched folders, so nested matches are removed together with their parent.

    Removal strategy (-Mode):
      - Auto (default): uses the Windows Recycle Bin on local fixed drives. On network drives
        (e.g. mapped X:), removable drives, and UNC paths, where Windows provides no Recycle Bin
        (files would be deleted permanently), the folder is moved into a dated trash folder instead:
            <TrashPath>\<yyyyMMdd_HHmmss>\<path relative to the volume root>
        The default TrashPath is "<volume root>\_DuplicatesTrash" (e.g. "X:\_DuplicatesTrash").
      - RecycleBin: always uses the Windows Recycle Bin (only allowed on local fixed drives).
      - TrashFolder: always moves folders into the dated trash folder.

    Safety:
      - ConfirmImpact is 'High', so PowerShell asks for confirmation for each folder unless
        -Force (or -Confirm:$false) is specified. -WhatIf only lists what would be removed.
      - Matched folders that are reparse points (junctions/symlinks) are skipped with a warning.
      - The trash folder itself is never scanned, even when -Path points to the volume root.
      - The scan root itself is never removed.

.PARAMETER Path
    Root directory to scan recursively. Defaults to "X:\Photos".

.PARAMETER DuplicateFolderName
    Name of the folders to remove. Defaults to "_Duplicates" (same default as Find-MediaDuplicates.ps1).

.PARAMETER Mode
    Removal strategy: Auto (default), RecycleBin, or TrashFolder.

.PARAMETER TrashPath
    Root of the trash folder used by TrashFolder mode. Defaults to "<volume root>\_DuplicatesTrash".
    Must be located on the same volume root as -Path (folders are moved by fast rename).

.PARAMETER Force
    Suppresses per-folder confirmation prompts.

.PARAMETER LogPath
    Optional path to the log file. Defaults to "RemoveDuplicateFolders_<timestamp>.log" in the script root.

.PARAMETER PassThru
    If specified, returns a summary object with removal statistics.

.EXAMPLE
    .\Remove-DuplicateFolders.ps1 -WhatIf
    Lists all _Duplicates folders below X:\Photos with file counts and sizes, without removing anything.

.EXAMPLE
    .\Remove-DuplicateFolders.ps1 -Path "X:\Pictures" -Force
    Removes all _Duplicates folders below X:\Pictures without prompting (moved to X:\_DuplicatesTrash\<timestamp>).
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Path = "X:\Pictures",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DuplicateFolderName = "_Duplicates",

    [Parameter()]
    [ValidateSet("Auto", "RecycleBin", "TrashFolder")]
    [string]$Mode = "Auto",

    [Parameter()]
    [string]$TrashPath = "",

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [string]$LogPath = "",

    [Parameter()]
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 8,

    [Parameter()]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "SmartMediaCommon.ps1"
if (-not (Test-Path -LiteralPath $commonModulePath -PathType Leaf)) {
    throw "Required shared module not found: '$commonModulePath'"
}
. $commonModulePath

#region Helper Functions

function Test-IsReparsePoint {
    <#
    .SYNOPSIS
        Returns $true if the directory is a junction, symbolic link, or other reparse point.
    #>
    param([Parameter(Mandatory = $true)][System.IO.DirectoryInfo]$Directory)
    return [bool]($Directory.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
}

function Test-IsSameOrChildPath {
    <#
    .SYNOPSIS
        Returns $true if CandidatePath equals ParentPath or is located below it (case-insensitive).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$CandidatePath,
        [Parameter(Mandatory = $true)][string]$ParentPath
    )
    $candidate = (Get-NormalizedPath -Path $CandidatePath).TrimEnd('\')
    $parent = (Get-NormalizedPath -Path $ParentPath).TrimEnd('\')
    return (
        $candidate.Equals($parent, [System.StringComparison]::OrdinalIgnoreCase) -or
        $candidate.StartsWith($parent + "\", [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Get-VolumeDriveType {
    <#
    .SYNOPSIS
        Returns the System.IO.DriveType of the volume containing the path (UNC paths are reported as Network).
    #>
    param([Parameter(Mandatory = $true)][string]$TargetPath)
    $root = [System.IO.Path]::GetPathRoot($TargetPath)
    if ($root.StartsWith("\\")) {
        return [System.IO.DriveType]::Network
    }
    try {
        return ([System.IO.DriveInfo]::new($root)).DriveType
    }
    catch {
        return [System.IO.DriveType]::Unknown
    }
}

function Find-DuplicateFolders {
    <#
    .SYNOPSIS
        Returns top-most folders named FolderName below RootPath without following reparse points
        and without descending into matched folders or excluded paths.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$RootPath,
        [Parameter(Mandatory = $true)][string]$FolderName,
        [Parameter()][string[]]$ExcludedPaths = @(),
        [Parameter(Mandatory = $true)][string]$LogFile,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Stats
    )

    $results = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()
    $stack = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
    $stack.Push([System.IO.DirectoryInfo]::new($RootPath))

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            $children = $current.GetDirectories()
        }
        catch {
            $Stats.Errors++
            Write-SorterLog -Message "Cannot list directory '$($current.FullName)': $($_.Exception.Message)" -Level "WARN" -LogFile $LogFile
            continue
        }

        foreach ($child in $children) {
            $isExcluded = $false
            foreach ($excluded in $ExcludedPaths) {
                if (Test-IsSameOrChildPath -CandidatePath $child.FullName -ParentPath $excluded) {
                    $isExcluded = $true
                    break
                }
            }
            if ($isExcluded) {
                Write-SorterLog -Message "Skipping excluded directory '$($child.FullName)'" -Level "INFO" -LogFile $LogFile
                continue
            }

            if (Test-IsIgnoredDirectoryName -DirectoryName $child.Name) {
                Write-SorterLog -Message "Skipping ignored directory '$($child.FullName)'" -Level "INFO" -LogFile $LogFile
                continue
            }

            if ($child.Name.Equals($FolderName, [System.StringComparison]::OrdinalIgnoreCase)) {
                if (Test-IsReparsePoint -Directory $child) {
                    $Stats.FoldersSkipped++
                    Write-SorterLog -Message "Skipping '$($child.FullName)': folder is a junction/symbolic link (target data is left untouched)" -Level "WARN" -LogFile $LogFile
                }
                else {
                    $results.Add($child)
                }
                continue
            }

            if (-not (Test-IsReparsePoint -Directory $child)) {
                $stack.Push($child)
            }
        }
    }

    return , $results
}

function Measure-DirectoryContent {
    <#
    .SYNOPSIS
        Counts files and total bytes inside a directory tree without following reparse points.
    #>
    param([Parameter(Mandatory = $true)][System.IO.DirectoryInfo]$Directory)

    [long]$fileCount = 0
    [long]$byteCount = 0
    $stack = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
    $stack.Push($Directory)

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            foreach ($file in $current.GetFiles()) {
                $fileCount++
                $byteCount += $file.Length
            }
            foreach ($sub in $current.GetDirectories()) {
                if (-not (Test-IsReparsePoint -Directory $sub)) {
                    $stack.Push($sub)
                }
            }
        }
        catch {
            # Unreadable subfolders are ignored for statistics; removal will report real errors
            continue
        }
    }

    return [PSCustomObject]@{ Files = $fileCount; Bytes = $byteCount }
}

function Get-TrashTargetPath {
    <#
    .SYNOPSIS
        Builds a non-colliding destination path inside the trash session folder that mirrors
        the source folder's path relative to its volume root.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$SourceFolder,
        [Parameter(Mandatory = $true)][string]$SessionFolder
    )
    $root = [System.IO.Path]::GetPathRoot($SourceFolder)
    $relative = $SourceFolder.Substring($root.Length).TrimStart('\')
    $target = Join-Path -Path $SessionFolder -ChildPath $relative

    $candidate = $target
    $index = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = "{0}_{1:D3}" -f $target, $index
        $index++
    }
    return $candidate
}

#endregion

#region Main

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path $PSScriptRoot -ChildPath "RemoveDuplicateFolders_$timestamp.log"
}

# -Force suppresses confirmation prompts unless -Confirm was explicitly requested
if ($Force -and -not $PSBoundParameters.ContainsKey('Confirm')) {
    $ConfirmPreference = 'None'
}

$stats = [ordered]@{
    Path             = ""
    Mode             = ""
    TrashSession     = ""
    FoldersFound     = 0
    FoldersRemoved   = 0
    FoldersSkipped   = 0
    FilesRemoved     = [long]0
    BytesRemoved     = [long]0
    Errors           = 0
}

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    throw "Directory not found: '$Path'"
}
$rootPath = Get-NormalizedPath -Path $Path
$stats.Path = $rootPath

if ((Split-Path -Path $rootPath -Leaf).Equals($DuplicateFolderName, [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-SorterLog -Message "Refusing to run: scan root '$rootPath' is itself a '$DuplicateFolderName' folder. Pass its parent directory instead." -Level "ERROR" -LogFile $LogPath
    throw "Scan root '$rootPath' is itself a '$DuplicateFolderName' folder."
}

if (Test-IsRecycleBinOrTrashPath -Path $rootPath) {
    Write-SorterLog -Message "Refusing to run: scan root '$rootPath' is inside a Recycle Bin or trash directory." -Level "ERROR" -LogFile $LogPath
    throw "Scan root '$rootPath' is inside a Recycle Bin or trash directory."
}

# Resolve effective removal mode
$volumeRoot = [System.IO.Path]::GetPathRoot($rootPath)
$driveType = Get-VolumeDriveType -TargetPath $rootPath
$hasRecycleBin = ($driveType -eq [System.IO.DriveType]::Fixed)

$effectiveMode = switch ($Mode) {
    "Auto" { if ($hasRecycleBin) { "RecycleBin" } else { "TrashFolder" } }
    default { $Mode }
}

if ($effectiveMode -eq "RecycleBin" -and -not $hasRecycleBin) {
    $msg = "Windows Recycle Bin is not available on '$volumeRoot' (DriveType=$driveType); files would be deleted permanently. Use -Mode Auto or -Mode TrashFolder."
    Write-SorterLog -Message $msg -Level "ERROR" -LogFile $LogPath
    throw $msg
}

if ([string]::IsNullOrWhiteSpace($TrashPath)) {
    $TrashPath = Join-Path -Path $volumeRoot -ChildPath "_DuplicatesTrash"
}
$TrashPath = Get-NormalizedPath -Path $TrashPath

$sessionFolder = ""
if ($effectiveMode -eq "TrashFolder") {
    $trashRoot = [System.IO.Path]::GetPathRoot($TrashPath)
    if (-not $trashRoot.Equals($volumeRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $msg = "TrashPath '$TrashPath' must be on the same volume root as '$rootPath' ($volumeRoot)."
        Write-SorterLog -Message $msg -Level "ERROR" -LogFile $LogPath
        throw $msg
    }
    if (Test-IsSameOrChildPath -CandidatePath $TrashPath -ParentPath (Join-Path -Path $rootPath -ChildPath $DuplicateFolderName)) {
        throw "TrashPath '$TrashPath' must not be located inside a '$DuplicateFolderName' folder."
    }
    $sessionFolder = Join-Path -Path $TrashPath -ChildPath $timestamp
    $stats.TrashSession = $sessionFolder
}
else {
    Add-Type -AssemblyName Microsoft.VisualBasic
}
$stats.Mode = $effectiveMode

Write-SorterLog -Message "Remove-DuplicateFolders started. Path='$rootPath', Folder='$DuplicateFolderName', Mode=$effectiveMode (requested $Mode, DriveType=$driveType), ThrottleLimit=$ThrottleLimit$(if ($sessionFolder) { ", TrashSession='$sessionFolder'" })" -Level "INFO" -LogFile $LogPath
Write-Host "Scanning '$rootPath' for '$DuplicateFolderName' folders (mode: $effectiveMode, ThrottleLimit=$ThrottleLimit)..." -ForegroundColor Cyan

$excludedPathsList = @($TrashPath)
$measuredTargets = [System.Collections.Generic.List[object]]::new()

# Inspect top-level children of rootPath first so each top-level subdirectory can be scanned in its own thread
$topLevelSubDirs = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()
try {
    $rootChildren = ([System.IO.DirectoryInfo]::new($rootPath)).GetDirectories()
}
catch {
    $stats.Errors++
    Write-SorterLog -Message "Cannot list directory '$rootPath': $($_.Exception.Message)" -Level "WARN" -LogFile $LogPath
    $rootChildren = @()
}

foreach ($child in $rootChildren) {
    $isExcluded = $false
    foreach ($excluded in $excludedPathsList) {
        if (Test-IsSameOrChildPath -CandidatePath $child.FullName -ParentPath $excluded) {
            $isExcluded = $true
            break
        }
    }
    if ($isExcluded) {
        Write-SorterLog -Message "Skipping excluded directory '$($child.FullName)'" -Level "INFO" -LogFile $LogPath
        continue
    }

    if (Test-IsIgnoredDirectoryName -DirectoryName $child.Name) {
        Write-SorterLog -Message "Skipping ignored directory '$($child.FullName)'" -Level "INFO" -LogFile $LogPath
        continue
    }

    if ($child.Name.Equals($DuplicateFolderName, [System.StringComparison]::OrdinalIgnoreCase)) {
        if (Test-IsReparsePoint -Directory $child) {
            $stats.FoldersSkipped++
            Write-SorterLog -Message "Skipping '$($child.FullName)': folder is a junction/symbolic link (target data is left untouched)" -Level "WARN" -LogFile $LogPath
        }
        else {
            $content = Measure-DirectoryContent -Directory $child
            $measuredTargets.Add([PSCustomObject]@{
                Directory = $child
                Files     = $content.Files
                Bytes     = $content.Bytes
            })
        }
        continue
    }

    if (-not (Test-IsReparsePoint -Directory $child)) {
        $topLevelSubDirs.Add($child)
    }
}

if ($ThrottleLimit -gt 1 -and $topLevelSubDirs.Count -gt 1) {
    $totalSub = $topLevelSubDirs.Count
    $doneSub = 0

    $topLevelSubDirs | ForEach-Object -Parallel {
        $subRootDir = $_
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        . $using:commonModulePath

        $folderName = $using:DuplicateFolderName
        $excludedList = [string[]]$using:excludedPathsList

        $foundList = [System.Collections.Generic.List[object]]::new()
        $warnList = [System.Collections.Generic.List[string]]::new()
        $skippedCount = 0
        $errCount = 0

        $stack = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
        $stack.Push($subRootDir)

        while ($stack.Count -gt 0) {
            $curr = $stack.Pop()
            try {
                $children = $curr.GetDirectories()
            }
            catch {
                $errCount++
                $warnList.Add("Cannot list directory '$($curr.FullName)': $($_.Exception.Message)")
                continue
            }

            foreach ($ch in $children) {
                $candNorm = (Get-NormalizedPath -Path $ch.FullName).TrimEnd('\')
                $isExcl = $false
                foreach ($ex in $excludedList) {
                    $exNorm = (Get-NormalizedPath -Path $ex).TrimEnd('\')
                    if ($candNorm.Equals($exNorm, [System.StringComparison]::OrdinalIgnoreCase) -or
                        $candNorm.StartsWith($exNorm + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
                        $isExcl = $true
                        break
                    }
                }
                if ($isExcl) { continue }
                if (Test-IsIgnoredDirectoryName -DirectoryName $ch.Name) { continue }

                $isReparse = [bool]($ch.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
                if ($ch.Name.Equals($folderName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    if ($isReparse) {
                        $skippedCount++
                        $warnList.Add("Skipping '$($ch.FullName)': folder is a junction/symbolic link (target data is left untouched)")
                    }
                    else {
                        [long]$fCnt = 0
                        [long]$bCnt = 0
                        $mStack = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
                        $mStack.Push($ch)
                        while ($mStack.Count -gt 0) {
                            $mCurr = $mStack.Pop()
                            try {
                                foreach ($f in $mCurr.GetFiles()) {
                                    $fCnt++
                                    $bCnt += $f.Length
                                }
                                foreach ($s in $mCurr.GetDirectories()) {
                                    if (-not [bool]($s.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                                        $mStack.Push($s)
                                    }
                                }
                            }
                            catch { continue }
                        }
                        $foundList.Add([PSCustomObject]@{
                            Directory = $ch
                            Files     = $fCnt
                            Bytes     = $bCnt
                        })
                    }
                    continue
                }

                if (-not $isReparse) {
                    $stack.Push($ch)
                }
            }
        }

        $sw.Stop()
        [PSCustomObject]@{
            SubRoot      = $subRootDir.FullName
            Targets      = $foundList.ToArray()
            Warnings     = $warnList.ToArray()
            SkippedCount = $skippedCount
            ErrorCount   = $errCount
            ElapsedMs    = $sw.ElapsedMilliseconds
        }
    } -ThrottleLimit $ThrottleLimit | ForEach-Object {
        $sr = $_
        $doneSub++
        $stats.FoldersSkipped += $sr.SkippedCount
        $stats.Errors += $sr.ErrorCount
        foreach ($w in $sr.Warnings) {
            Write-SorterLog -Message $w -Level "WARN" -LogFile $LogPath
        }
        foreach ($t in $sr.Targets) {
            $measuredTargets.Add($t)
        }
        Write-Host ("[{0}/{1}] Scanned '{2}' (found {3} '{4}' folder(s), {5} ms)" -f $doneSub, $totalSub, $sr.SubRoot, $sr.Targets.Count, $DuplicateFolderName, $sr.ElapsedMs) -ForegroundColor DarkCyan
    }
}
else {
    foreach ($subDir in $topLevelSubDirs) {
        $subFound = Find-DuplicateFolders -RootPath $subDir.FullName -FolderName $DuplicateFolderName -ExcludedPaths $excludedPathsList -LogFile $LogPath -Stats $stats
        foreach ($sf in $subFound) {
            $content = Measure-DirectoryContent -Directory $sf
            $measuredTargets.Add([PSCustomObject]@{
                Directory = $sf
                Files     = $content.Files
                Bytes     = $content.Bytes
            })
        }
    }
}

$stats.FoldersFound = $measuredTargets.Count
Write-SorterLog -Message "Found $($measuredTargets.Count) '$DuplicateFolderName' folder(s)." -Level "INFO" -LogFile $LogPath

foreach ($item in $measuredTargets) {
    $folder = $item.Directory
    $content = [PSCustomObject]@{ Files = $item.Files; Bytes = $item.Bytes }
    $sizeMb = [math]::Round($content.Bytes / 1MB, 2)
    $description = "$($folder.FullName) ($($content.Files) files, $sizeMb MB)"

    $trashTarget = ""
    if ($effectiveMode -eq "TrashFolder") {
        $trashTarget = Get-TrashTargetPath -SourceFolder $folder.FullName -SessionFolder $sessionFolder
        $action = "Move to trash folder '$trashTarget'"
    }
    else {
        $action = "Send to Recycle Bin"
    }

    if (-not $PSCmdlet.ShouldProcess($description, $action)) {
        Write-SorterLog -Message "WhatIf: Would remove [$effectiveMode] $description$(if ($trashTarget) { " -> '$trashTarget'" })" -Level "WHATIF" -LogFile $LogPath
        continue
    }

    try {
        if ($effectiveMode -eq "TrashFolder") {
            $parent = Split-Path -Path $trashTarget -Parent
            if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
                New-Item -Path $parent -ItemType Directory -Force -WhatIf:$false | Out-Null
            }
            [System.IO.Directory]::Move($folder.FullName, $trashTarget)
            Write-SorterLog -Message "Moved $description -> '$trashTarget'" -Level "MOVED" -LogFile $LogPath
        }
        else {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                $folder.FullName,
                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
            )
            Write-SorterLog -Message "Sent to Recycle Bin: $description" -Level "DELETED" -LogFile $LogPath
        }

        $stats.FoldersRemoved++
        $stats.FilesRemoved += $content.Files
        $stats.BytesRemoved += $content.Bytes
    }
    catch {
        $stats.Errors++
        Write-SorterLog -Message "Failed to remove '$($folder.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $LogPath
    }
}

$freedMb = [math]::Round($stats.BytesRemoved / 1MB, 2)
$summary = "Summary: Found=$($stats.FoldersFound), Removed=$($stats.FoldersRemoved), Skipped=$($stats.FoldersSkipped), Files=$($stats.FilesRemoved), Size=$freedMb MB, Errors=$($stats.Errors)"
Write-SorterLog -Message $summary -Level "INFO" -LogFile $LogPath
Write-Host $summary -ForegroundColor $(if ($stats.Errors -gt 0) { "Yellow" } else { "Green" })
if ($sessionFolder -and $stats.FoldersRemoved -gt 0) {
    Write-Host "Removed folders were moved to '$sessionFolder'. Delete it manually once verified." -ForegroundColor Yellow
}
Write-Host "Log file: $LogPath"

if ($PassThru) {
    return [PSCustomObject]$stats
}

#endregion
