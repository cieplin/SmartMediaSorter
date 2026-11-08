#Requires -Version 7.0

<#
.SYNOPSIS
    Detects duplicate files using a multi-stage strategy (filename hint first with whole-file
    and metadata-stripped image payload SHA-256, then content SHA-256 fallback, plus reference
    library database check) and moves duplicates into a subfolder (default: _Duplicates)
    in the same directory.

.DESCRIPTION
    Find-MediaDuplicates scans a target directory (default: X:\Photos) and identifies duplicate
    files according to the following rules:
    1. System metadata files (Thumbs.db, desktop.ini, .DS_Store) and standalone sidecar files
       (.xmp, .on1, .dop, .pp3) are automatically excluded from independent duplicate scanning.
    2. Different extensions mean different files: Files with different extensions (e.g., .jpg vs .png)
       are never considered duplicates.
    3. Stage 1 (First hint - Filename normalization): Files with the same extension are grouped by their
       normalized base name (Get-NormalizedBaseName), which groups originals with common copy/burst/revision
       naming patterns (e.g., "IMG_0001.jpg", "IMG_0001 (1).jpg", "IMG_0001 - kopia.jpg",
       "2025-01-02_153726_000.jpg", "2025-01-02_153726-01.jpg", "2025-01-02_153731.jpg").
       - Step 1A: Candidates with identical byte length are compared by whole-file SHA-256.
       - Step 1B: Candidates in the same name group whose byte lengths differ (e.g., due to edited
         EXIF/XMP metadata headers) are compared by metadata-stripped image payload SHA-256
         (Get-CachedImagePayloadHash), catching identical photos with different metadata while
         preserving genuine burst shots.
    4. Stage 2 (Second check - Content fallback within same extension): Remaining files that did not
       match by filename in Stage 1 are grouped by (Extension + Byte Length) and verified by SHA-256
       hash so identical files with completely different names are also caught.
    5. Stage 3 (Reference Library / Database check against X:\Pictures): Remaining files are compared
       by (Extension + Byte Length + SHA-256) against existing files in the reference destination
       database (.smartmediasorter_db.json in ReferencePath, default: X:\Pictures).
    6. Action - Keep best original & move duplicates to local subfolder:
       Original selection priority:
         a) File that has an associated sidecar file (.xmp, .on1, .dop, .pp3)
         b) File with richer metadata (larger byte length when image payload is identical)
         c) Clean filename without OS copy suffixes (" (1)", " - kopia", "_copy")
         d) Shortest filename / alphabetical order.
       All remaining duplicates in the set (along with any duplicate-specific sidecars) are moved
       to "<DuplicateFolderName>" (default: "_Duplicates") in the same directory.

.PARAMETER Path
    The directory to scan for duplicate files. Defaults to "X:\Photos".

.PARAMETER ReferencePath
    Optional reference library directory (default: "X:\Pictures") whose hash database
    (.smartmediasorter_db.json) is used to detect duplicates of files already organized in the destination.

.PARAMETER DuplicateFolderName
    Name of the subfolder created inside the duplicate file's directory where duplicates
    will be moved. Defaults to "_Duplicates".

.PARAMETER Recurse
    Whether to scan subdirectories within Path. Defaults to $true. Existing folders matching
    DuplicateFolderName are automatically excluded from scanning.

.PARAMETER NameMatchOnly
    If specified, only Stage 1 (files with matching normalized base names) is executed.

.PARAMETER DatabasePath
    Optional path to the persistent JSON hash database (.smartmediasorter_db.json).
    Automatically uses "<ReferencePath>\.smartmediasorter_db.json" if present, otherwise
    "<Path>\.smartmediasorter_db.json".

.PARAMETER NoDatabase
    If specified, disables reading from and writing to the persistent hash database.

.PARAMETER RebuildDatabase
    If specified, ignores existing cached hashes and rebuilds database entries for scanned files.

.PARAMETER LogPath
    Optional path to the log file. Defaults to "FindMediaDuplicates_<timestamp>.log" in the script root.

.PARAMETER IncludeExtensions
    Optional list of file extensions (with leading dot) to include. When empty (default),
    all non-system, non-sidecar file extensions are scanned.

.PARAMETER PassThru
    If specified, returns a summary object with duplicate detection statistics.

.EXAMPLE
    .\Find-MediaDuplicates.ps1 -Path "X:\Pictures\2025" -WhatIf
    Simulates finding and moving duplicates in X:\Pictures\2025 (including metadata-divergent duplicates).

.EXAMPLE
    .\Find-MediaDuplicates.ps1 -Path "X:\Pictures\2025"
    Scans X:\Pictures\2025 and moves duplicates to local _Duplicates folders.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Path = "X:\Pictures",

    [Parameter()]
    [string]$ReferencePath = "X:\Pictures",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DuplicateFolderName = "_Duplicates",

    [Parameter()]
    [bool]$Recurse = $true,

    [Parameter()]
    [switch]$NameMatchOnly,

    [Parameter()]
    [ValidateRange(0, 16)]
    [int]$MaxPerceptualDistance = 4,

    [Parameter()]
    [string]$DatabasePath = "",

    [Parameter()]
    [switch]$NoDatabase,

    [Parameter()]
    [switch]$RebuildDatabase,

    [Parameter()]
    [string]$LogPath = "",

    [Parameter()]
    [string[]]$IncludeExtensions = @(),

    [Parameter()]
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 8,

    [Parameter()]
    [ValidateRange(10, 3600)]
    [int]$DirectoryTimeoutSeconds = 300,

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

function Test-IsInsideDuplicateFolder {
    <#
    .SYNOPSIS
        Returns $true if the file's directory hierarchy contains a folder matching DuplicateFolderName.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FileDirectoryPath,
        [Parameter(Mandatory = $true)][string]$RootScanPath,
        [Parameter(Mandatory = $true)][string]$DupFolderName
    )

    $normDir = Get-NormalizedPath -Path $FileDirectoryPath
    $normRoot = Get-NormalizedPath -Path $RootScanPath

    if ($normDir.Equals($normRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $relative = if ($normDir.StartsWith($normRoot + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        $normDir.Substring($normRoot.Length + 1)
    }
    else {
        $normDir
    }

    $segments = $relative -split '[\\/]'
    foreach ($seg in $segments) {
        if ($seg.Equals($DupFolderName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Test-HasOsCopyMarker {
    <#
    .SYNOPSIS
        Returns $true if a filename stem contains explicit OS copy markers such as " (1)", " - kopia", "_copy".
    #>
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$BaseName)
    if ([string]::IsNullOrWhiteSpace($BaseName)) {
        return $false
    }
    $copyPattern = '(?i)(\s*-\s*(copy|kopia)(\s*\(\d+\)|\s+\d+)?|\s*[_\-\s]+(copy|kopia)(\s*\(\d+\)|\s*_\d+|\s+\d+)?|\s*\(\d+\)|\s*\[\d+\])$'
    return ($BaseName -match $copyPattern)
}

function Move-SingleDuplicateToLocalSubfolder {
    <#
    .SYNOPSIS
        Moves a single confirmed duplicate file (and any exact-named sidecar file attached to it)
        into its local _Duplicates folder and updates database state.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$DuplicateFile,
        [Parameter(Mandatory = $true)][string]$OriginalFileFullPath,
        [Parameter(Mandatory = $true)][string]$OriginalFileName,
        [Parameter(Mandatory = $true)][string]$Sha256Hash,
        [Parameter(Mandatory = $true)][string]$DetectionStage,
        [Parameter(Mandatory = $true)][string]$DupFolderName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$PlannedTargets,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$QuarantinedFiles,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Database,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$SourceDirsToRefresh,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$DestDirsToRefresh,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Stats,
        [Parameter(Mandatory = $true)][string]$LogFile
    )

    $normDupPath = Get-NormalizedPath -Path $DuplicateFile.FullName
    [void]$QuarantinedFiles.Add($normDupPath)

    $normDupDir = Get-NormalizedPath -Path $DuplicateFile.DirectoryName
    $targetDupFolder = Join-Path -Path $normDupDir -ChildPath $DupFolderName
    $normTargetDupFolder = Get-NormalizedPath -Path $targetDupFolder

    # Determine non-colliding filename inside the local _Duplicates folder
    $dupBase = [System.IO.Path]::GetFileNameWithoutExtension($DuplicateFile.Name)
    $dupExt = $DuplicateFile.Extension
    $candidateTargetName = $DuplicateFile.Name
    $candidateTargetPath = Get-NormalizedPath -Path (Join-Path -Path $normTargetDupFolder -ChildPath $candidateTargetName)

    $collisionIdx = 1
    while ((Test-Path -LiteralPath $candidateTargetPath -PathType Leaf) -or $PlannedTargets.Contains($candidateTargetPath)) {
        $candidateTargetName = "{0}_{1:D3}{2}" -f $dupBase, $collisionIdx, $dupExt
        $candidateTargetPath = Get-NormalizedPath -Path (Join-Path -Path $normTargetDupFolder -ChildPath $candidateTargetName)
        $collisionIdx++
    }

    [void]$PlannedTargets.Add($candidateTargetPath)

    if ($PSCmdlet.ShouldProcess("$normDupPath -> $candidateTargetPath", "Move duplicate ($DetectionStage) of '$OriginalFileName'")) {
        if (-not (Test-Path -LiteralPath $normTargetDupFolder -PathType Container)) {
            New-Item -Path $normTargetDupFolder -ItemType Directory -Force | Out-Null
        }

        Move-Item -LiteralPath $normDupPath -Destination $candidateTargetPath -Force

        # If this duplicate file had an associated sidecar ("<FullName>.<ext>" or "<BaseName>.<ext>") that does NOT belong to the kept original, move it too
        $origBasePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($OriginalFileFullPath), [System.IO.Path]::GetFileNameWithoutExtension($OriginalFileFullPath))
        $dupBasePath = [System.IO.Path]::Combine($normDupDir, $dupBase)
        $targetBasePath = [System.IO.Path]::Combine($normTargetDupFolder, [System.IO.Path]::GetFileNameWithoutExtension($candidateTargetName))

        foreach ($scExt in @(".xmp", ".on1", ".dop", ".pp3")) {
            $dupExactSidecar = $normDupPath + $scExt
            if ([System.IO.File]::Exists($dupExactSidecar) -and
                (-not $dupExactSidecar.Equals($OriginalFileFullPath + $scExt, [System.StringComparison]::OrdinalIgnoreCase))) {
                $scTarget = $candidateTargetPath + $scExt
                Move-Item -LiteralPath $dupExactSidecar -Destination $scTarget -Force
                Write-SorterLog -Message "Moved associated sidecar '$dupExactSidecar' -> '$scTarget'" -Level "MOVED" -LogFile $LogFile
            }

            $dupStemSidecar = $dupBasePath + $scExt
            if ([System.IO.File]::Exists($dupStemSidecar) -and
                (-not $dupStemSidecar.Equals($origBasePath + $scExt, [System.StringComparison]::OrdinalIgnoreCase))) {
                $scStemTarget = $targetBasePath + $scExt
                Move-Item -LiteralPath $dupStemSidecar -Destination $scStemTarget -Force
                Write-SorterLog -Message "Moved associated sidecar '$dupStemSidecar' -> '$scStemTarget'" -Level "MOVED" -LogFile $LogFile
            }
        }

        # Remove moved duplicate from active database file tables and mark directory for hash refresh
        $Database["SourceFiles"].Remove($normDupPath)
        $Database["DestinationFiles"].Remove($normDupPath)

        $isDestDir = $Database["DestinationDirectories"].ContainsKey($normDupDir)
        if (-not $isDestDir -and -not [string]::IsNullOrWhiteSpace($ReferencePath)) {
            $normRefRoot = Get-NormalizedPath -Path $ReferencePath
            $isDestDir = (
                $normDupDir.Equals($normRefRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
                $normDupDir.StartsWith($normRefRoot + "\", [System.StringComparison]::OrdinalIgnoreCase)
            )
        }

        if ($isDestDir) {
            [void]$DestDirsToRefresh.Add($normDupDir)
        }
        else {
            [void]$SourceDirsToRefresh.Add($normDupDir)
        }

        $Stats.DuplicatesMoved++
        Write-SorterLog -Message "Moved duplicate [$DetectionStage] '$normDupPath' -> '$candidateTargetPath' (Matches original '$OriginalFileFullPath', Hash=$($Sha256Hash.Substring(0, 12))...)" -Level "MOVED" -LogFile $LogFile
    }
    else {
        Write-SorterLog -Message "WhatIf: Would move duplicate [$DetectionStage] '$normDupPath' -> '$candidateTargetPath' (Matches original '$OriginalFileFullPath', Hash=$($Sha256Hash.Substring(0, 12))...)" -Level "WHATIF" -LogFile $LogFile
    }
}

function Move-ConfirmedDuplicateGroup {
    <#
    .SYNOPSIS
        Selects the primary (original) file to keep from a confirmed duplicate group and moves
        all remaining duplicate copies into their local _Duplicates subfolders.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[System.IO.FileInfo]]$IdenticalFiles,
        [Parameter(Mandatory = $true)][string]$Sha256Hash,
        [Parameter(Mandatory = $true)][string]$DetectionStage,
        [Parameter()][string]$PreferredStem = "",
        [Parameter(Mandatory = $true)][string]$DupFolderName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$PlannedTargets,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$QuarantinedFiles,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$KeptOriginals,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Database,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$SourceDirsToRefresh,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$DestDirsToRefresh,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Stats,
        [Parameter(Mandatory = $true)][string]$LogFile
    )

    # Filter out any files that were already moved in an earlier sub-step
    $activeFiles = @(
        $IdenticalFiles | Where-Object {
            -not $QuarantinedFiles.Contains((Get-NormalizedPath -Path $_.FullName))
        }
    )
    if ($activeFiles.Count -lt 2) {
        return
    }

    $Stats.DuplicateSetsFound++

    $maxGroupLength = [long](($activeFiles | Measure-Object -Property Length -Maximum).Maximum)
    $fullQualityThreshold = [long]($maxGroupLength * 0.85)

    # Select the primary ("Original") file to keep:
    # 1. Full-quality tier (>= 85% of max byte Length in the group) -> 0 vs 1 (prevents small thumbnails from beating full-res images)
    # 2. File that has an associated sidecar file (.xmp, .on1, .dop, .pp3) -> 0 vs 1
    # 3. File with richer metadata / higher quality (larger byte Length descending)
    # 4. File without OS copy markers (" (1)", " - kopia", "_copy") -> 0 vs 1
    # 5. File already selected as a kept original in an earlier pass -> 0 vs 1
    # 6. File matching standard SmartMediaSorter _000 suffix or exact normalized stem -> 0 vs 1
    # 7. Shortest file name (e.g. "DSC_5000.jpg" before "RENAMED_BACKUP_99.jpg")
    # 8. Alphabetical FullName
    $sortedIdentical = @(
        $activeFiles | Sort-Object -Property @(
            @{ Expression = {
                if ($_.Length -ge $fullQualityThreshold) { 0 } else { 1 }
            }; Ascending = $true },
            @{ Expression = {
                if (Test-HasSidecarFile -FileItem $_) { 0 } else { 1 }
            }; Ascending = $true },
            @{ Expression = { $_.Length }; Ascending = $false },
            @{ Expression = {
                $stem = [System.IO.Path]::GetFileNameWithoutExtension($_.Name).Trim()
                if (Test-HasOsCopyMarker -BaseName $stem) { 1 } else { 0 }
            }; Ascending = $true },
            @{ Expression = {
                $normFull = Get-NormalizedPath -Path $_.FullName
                if ($KeptOriginals.Contains($normFull)) { 0 } else { 1 }
            }; Ascending = $true },
            @{ Expression = {
                $stem = [System.IO.Path]::GetFileNameWithoutExtension($_.Name).Trim().ToLowerInvariant()
                if ($stem -match '^\d{4}-\d{2}-\d{2}_\d{6}_000$' -or
                    (-not [string]::IsNullOrWhiteSpace($PreferredStem) -and $stem -eq $PreferredStem)) { 0 } else { 1 }
            }; Ascending = $true },
            @{ Expression = { $_.Name.Length }; Ascending = $true },
            @{ Expression = { $_.FullName }; Ascending = $true }
        )
    )

    $originalFile = $sortedIdentical[0]
    [void]$KeptOriginals.Add((Get-NormalizedPath -Path $originalFile.FullName))

    $duplicateFiles = @($sortedIdentical | Select-Object -Skip 1)

    foreach ($dupFile in $duplicateFiles) {
        try {
            Move-SingleDuplicateToLocalSubfolder `
                -DuplicateFile $dupFile `
                -OriginalFileFullPath $originalFile.FullName `
                -OriginalFileName $originalFile.Name `
                -Sha256Hash $Sha256Hash `
                -DetectionStage $DetectionStage `
                -DupFolderName $DupFolderName `
                -PlannedTargets $PlannedTargets `
                -QuarantinedFiles $QuarantinedFiles `
                -Database $Database `
                -SourceDirsToRefresh $SourceDirsToRefresh `
                -DestDirsToRefresh $DestDirsToRefresh `
                -Stats $Stats `
                -LogFile $LogFile `
                -WhatIf:$WhatIfPreference
        }
        catch {
            $Stats.Errors++
            Write-SorterLog -Message "Error moving duplicate '$($dupFile.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $LogFile
        }
    }
}

#endregion Helper Functions

#region Main Execution

# Initialize log file path
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $runTimestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $LogPath = Join-Path -Path $scriptRoot -ChildPath "FindMediaDuplicates_$runTimestamp.log"
}

$logDir = Split-Path -Path $LogPath -Parent
if ($logDir -and -not (Test-Path -LiteralPath $logDir -PathType Container)) {
    New-Item -Path $logDir -ItemType Directory -Force -WhatIf:$false | Out-Null
}
if (-not (Test-Path -LiteralPath $LogPath -PathType Leaf)) {
    New-Item -Path $LogPath -ItemType File -Force -WhatIf:$false | Out-Null
}

$resolvedLogPath = Get-NormalizedPath -Path (Resolve-Path -LiteralPath $LogPath).Path

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-SorterLog -Message "Target directory does not exist: '$Path'" -Level "ERROR" -LogFile $resolvedLogPath
    throw "Target directory does not exist: '$Path'"
}

$resolvedRoot = Get-NormalizedPath -Path (Resolve-Path -LiteralPath $Path).Path
$normalizedExtensions = @($IncludeExtensions | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLowerInvariant() })

if (Test-IsRecycleBinOrTrashPath -Path $resolvedRoot) {
    Write-SorterLog -Message "Refusing to run: scan root '$resolvedRoot' is inside a Recycle Bin or trash directory." -Level "ERROR" -LogFile $resolvedLogPath
    throw "Refusing to run: scan root '$resolvedRoot' is inside a Recycle Bin or trash directory."
}
if (-not [string]::IsNullOrWhiteSpace($ReferencePath) -and (Test-IsRecycleBinOrTrashPath -Path $ReferencePath)) {
    Write-SorterLog -Message "Refusing to run: reference path '$ReferencePath' is inside a Recycle Bin or trash directory." -Level "ERROR" -LogFile $resolvedLogPath
    throw "Refusing to run: reference path '$ReferencePath' is inside a Recycle Bin or trash directory."
}

$excludedSystemFileNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($sysName in @("Thumbs.db", "desktop.ini", ".DS_Store", ".smartmediasorter_db.json")) {
    [void]$excludedSystemFileNames.Add($sysName)
}

$excludedSidecarExtensions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($scExt in @(".xmp", ".on1", ".dop", ".pp3")) {
    [void]$excludedSidecarExtensions.Add($scExt)
}

# Resolve DatabasePath: prefer shared database in ReferencePath (e.g. X:\Pictures\.smartmediasorter_db.json) if it exists
if (-not $NoDatabase -and [string]::IsNullOrWhiteSpace($DatabasePath)) {
    $refDbCandidate = if (-not [string]::IsNullOrWhiteSpace($ReferencePath)) {
        Get-NormalizedPath -Path ([System.IO.Path]::Combine($ReferencePath, ".smartmediasorter_db.json"))
    } else {
        $null
    }

    if ($refDbCandidate -and [System.IO.File]::Exists($refDbCandidate)) {
        $DatabasePath = $refDbCandidate
    }
    else {
        $DatabasePath = [System.IO.Path]::Combine($resolvedRoot, ".smartmediasorter_db.json")
    }
}
elseif (-not [string]::IsNullOrWhiteSpace($DatabasePath)) {
    $DatabasePath = Get-NormalizedPath -Path $DatabasePath
}

Write-SorterLog -Message "Starting Find-MediaDuplicates (Path='$resolvedRoot', ReferencePath='$ReferencePath', DatabasePath='$DatabasePath', DuplicateFolder='$DuplicateFolderName', Recurse=$Recurse, ThrottleLimit=$ThrottleLimit, NameMatchOnly=$NameMatchOnly, WhatIf=$WhatIfPreference, UseDatabase=$(-not $NoDatabase))" -Level "INFO" -LogFile $resolvedLogPath

# Load or initialize shared hash database
$sorterDb = if (-not $NoDatabase -and -not $RebuildDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath)) {
    Import-SorterDatabase -DbFilePath $DatabasePath -LogFile $resolvedLogPath
}
else {
    New-EmptySorterDatabase
}

# Build list of directories to scan (skipping _Duplicates, _DuplicatesTrash, and reparse points)
$scanDirectories = [System.Collections.Generic.List[string]]::new()
$dirStack = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
$dirStack.Push([System.IO.DirectoryInfo]::new($resolvedRoot))

while ($dirStack.Count -gt 0) {
    $currDir = $dirStack.Pop()
    if (Test-IsIgnoredDirectoryName -DirectoryName $currDir.Name) {
        continue
    }
    $scanDirectories.Add((Get-NormalizedPath -Path $currDir.FullName))

    if ($Recurse) {
        try {
            $subDirs = $currDir.GetDirectories()
        }
        catch {
            Write-SorterLog -Message "Cannot list directory '$($currDir.FullName)': $($_.Exception.Message)" -Level "WARN" -LogFile $resolvedLogPath
            continue
        }

        for ($idx = $subDirs.Length - 1; $idx -ge 0; $idx--) {
            $sub = $subDirs[$idx]
            if ($sub.Name.Equals($DuplicateFolderName, [System.StringComparison]::OrdinalIgnoreCase) -or
                (Test-IsIgnoredDirectoryName -DirectoryName $sub.Name)) {
                continue
            }
            if ([bool]($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                continue
            }
            $dirStack.Push($sub)
        }
    }
}

$hashCache = @{}
$payloadHashCache = @{}
$perceptualHashCache = @{}

# Partition cached database file records by directory so each parallel worker only receives its own directory's records
$dbByDir = @{}
if (-not $NoDatabase) {
    foreach ($tableName in @("DestinationFiles", "SourceFiles")) {
        foreach ($kv in $sorterDb[$tableName].GetEnumerator()) {
            $rec = $kv.Value
            $dPath = if ($rec.ContainsKey("DirectoryPath") -and -not [string]::IsNullOrWhiteSpace($rec["DirectoryPath"])) {
                [string]$rec["DirectoryPath"]
            }
            else {
                [System.IO.Path]::GetDirectoryName([string]$kv.Key)
            }
            if (-not $dbByDir.ContainsKey($dPath)) {
                $dbByDir[$dPath] = @{}
            }
            $dbByDir[$dPath][$kv.Key] = $rec
        }
    }
}

$excludedSysList = @($excludedSystemFileNames)
$excludedScList = @($excludedSidecarExtensions)
$allScannedFilesList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$scanErrorsCount = 0

if ($ThrottleLimit -gt 1 -and $scanDirectories.Count -gt 1) {
    Write-Host "Scanning $($scanDirectories.Count) directories in parallel (ThrottleLimit=$ThrottleLimit, Timeout=${DirectoryTimeoutSeconds}s)..." -ForegroundColor Cyan
    $totalDirs = $scanDirectories.Count
    $completedDirs = 0
    $unsavedPreHashes = 0

    $workerScript = {
        param(
            [string]$dirPath,
            [hashtable]$localDbSnap,
            [string]$commonModulePath,
            [string[]]$excludedSysList,
            [string[]]$excludedScList,
            [string[]]$normalizedExtensions,
            [string]$resolvedLogPath,
            [string]$DatabasePath,
            [bool]$NameMatchOnly,
            [hashtable]$statusSync
        )

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        . $commonModulePath

        $localExcludedSys = [System.Collections.Generic.HashSet[string]]::new($excludedSysList, [System.StringComparer]::OrdinalIgnoreCase)
        $localExcludedSc = [System.Collections.Generic.HashSet[string]]::new($excludedScList, [System.StringComparer]::OrdinalIgnoreCase)

        $dirFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        $compSha256 = @{}
        $compPayload = @{}
        $compPerceptual = @{}
        $errMsg = $null

        try {
            $statusSync["Stage"] = "Listing files"
            $rawFiles = [System.IO.DirectoryInfo]::new($dirPath).GetFiles()
            foreach ($f in $rawFiles) {
                if ([bool]($f.Attributes -band ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)) -or
                    $f.Name.StartsWith('.')) {
                    continue
                }
                $normFull = Get-NormalizedPath -Path $f.FullName
                $extLow = $f.Extension.ToLowerInvariant()
                if (($normFull -ne $resolvedLogPath) -and
                    ([string]::IsNullOrWhiteSpace($DatabasePath) -or ($normFull -ne $DatabasePath)) -and
                    (-not $localExcludedSys.Contains($f.Name)) -and
                    (-not $localExcludedSc.Contains($extLow)) -and
                    ($normalizedExtensions.Count -eq 0 -or ($normalizedExtensions -contains $extLow))) {
                    $dirFiles.Add($f)
                }
            }

            $statusSync["FileTotal"] = $dirFiles.Count

            if ($dirFiles.Count -ge 2) {
                $localNameGroups = @{}
                foreach ($f in $dirFiles) {
                    $extLow = $f.Extension.ToLowerInvariant()
                    $stem = Get-NormalizedBaseName -BaseName ([System.IO.Path]::GetFileNameWithoutExtension($f.Name))
                    $gKey = "{0}|{1}" -f $extLow, $stem
                    if (-not $localNameGroups.ContainsKey($gKey)) {
                        $localNameGroups[$gKey] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                    }
                    $localNameGroups[$gKey].Add($f)
                }

                $locallyMatched = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

                foreach ($gKey in $localNameGroups.Keys) {
                    $cands = $localNameGroups[$gKey]
                    if ($cands.Count -lt 2) { continue }

                    $parts = $gKey -split '\|', 2
                    $extGrp = $parts[0]

                    # Step 1A pre-hash: same size in name group
                    $bySize = @{}
                    foreach ($c in $cands) {
                        $sk = [string]$c.Length
                        if (-not $bySize.ContainsKey($sk)) {
                            $bySize[$sk] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                        }
                        $bySize[$sk].Add($c)
                    }

                    foreach ($sk in $bySize.Keys) {
                        $sameSz = $bySize[$sk]
                        if ($sameSz.Count -lt 2) { continue }
                        $shaSeen = @{}
                        foreach ($fi in $sameSz) {
                            $nf = Get-NormalizedPath -Path $fi.FullName
                            $cachedSha = $null
                            if ($localDbSnap.ContainsKey($nf)) {
                                $rec = $localDbSnap[$nf]
                                if ($rec["Length"] -eq $fi.Length -and $rec["LastWriteTimeUtcTicks"] -eq $fi.LastWriteTimeUtc.Ticks -and -not [string]::IsNullOrWhiteSpace($rec["Sha256"])) {
                                    $cachedSha = [string]$rec["Sha256"]
                                }
                            }
                            if ([string]::IsNullOrWhiteSpace($cachedSha)) {
                                $statusSync["Stage"] = "SHA256 (1A)"
                                $statusSync["CurrentFile"] = $fi.Name
                                $cachedSha = (Get-FileHash -LiteralPath $nf -Algorithm SHA256).Hash
                                $compSha256[$nf] = $cachedSha
                            }
                            if ($shaSeen.ContainsKey($cachedSha)) {
                                [void]$locallyMatched.Add($nf)
                            }
                            else {
                                $shaSeen[$cachedSha] = $nf
                            }
                        }
                    }

                    # Step 1B & 1C pre-hash for image/RAW formats
                    if ($extGrp -in @(".jpg", ".jpeg", ".webp", ".png", ".dng", ".tif", ".tiff", ".cr2", ".nef", ".arw", ".orf", ".pef")) {
                        $rem1B = @($cands | Where-Object { -not $locallyMatched.Contains((Get-NormalizedPath -Path $_.FullName)) })
                        if ($rem1B.Count -ge 2) {
                            $payloadSeen = @{}
                            foreach ($fi in $rem1B) {
                                $nf = Get-NormalizedPath -Path $fi.FullName
                                $cachedPay = $null
                                if ($localDbSnap.ContainsKey($nf)) {
                                    $rec = $localDbSnap[$nf]
                                    if ($rec["Length"] -eq $fi.Length -and $rec["LastWriteTimeUtcTicks"] -eq $fi.LastWriteTimeUtc.Ticks -and -not [string]::IsNullOrWhiteSpace($rec["PayloadSha256"])) {
                                        $cachedPay = [string]$rec["PayloadSha256"]
                                    }
                                }
                                if ([string]::IsNullOrWhiteSpace($cachedPay)) {
                                    $statusSync["Stage"] = "PayloadHash (1B)"
                                    $statusSync["CurrentFile"] = $fi.Name
                                    $cachedPay = Get-ImagePayloadSha256 -FilePath $nf
                                    if (-not [string]::IsNullOrWhiteSpace($cachedPay)) {
                                        $compPayload[$nf] = $cachedPay
                                    }
                                }
                                if (-not [string]::IsNullOrWhiteSpace($cachedPay)) {
                                    if ($payloadSeen.ContainsKey($cachedPay)) {
                                        [void]$locallyMatched.Add($nf)
                                    }
                                    else {
                                        $payloadSeen[$cachedPay] = $nf
                                    }
                                }
                            }
                        }

                        $rem1C = @($cands | Where-Object { -not $locallyMatched.Contains((Get-NormalizedPath -Path $_.FullName)) })
                        if ($rem1C.Count -ge 2) {
                            foreach ($fi in $rem1C) {
                                $nf = Get-NormalizedPath -Path $fi.FullName
                                $cachedPer = $null
                                if ($localDbSnap.ContainsKey($nf)) {
                                    $rec = $localDbSnap[$nf]
                                    if ($rec["Length"] -eq $fi.Length -and $rec["LastWriteTimeUtcTicks"] -eq $fi.LastWriteTimeUtc.Ticks -and -not [string]::IsNullOrWhiteSpace($rec["PerceptualHash"])) {
                                        $cachedPer = [string]$rec["PerceptualHash"]
                                    }
                                }
                                if ([string]::IsNullOrWhiteSpace($cachedPer)) {
                                    $statusSync["Stage"] = "PerceptualHash (1C)"
                                    $statusSync["CurrentFile"] = $fi.Name
                                    $cachedPer = Get-ImagePerceptualHash -FilePath $nf
                                    if (-not [string]::IsNullOrWhiteSpace($cachedPer)) {
                                        $compPerceptual[$nf] = $cachedPer
                                    }
                                }
                            }
                        }
                    }
                }

                # Stage 2 intra-directory same-extension & same-size pre-hash
                if (-not $NameMatchOnly) {
                    $extSzGroups = @{}
                    foreach ($f in $dirFiles) {
                        $nf = Get-NormalizedPath -Path $f.FullName
                        if ($locallyMatched.Contains($nf)) { continue }
                        $esk = "{0}|{1}" -f $f.Extension.ToLowerInvariant(), $f.Length
                        if (-not $extSzGroups.ContainsKey($esk)) {
                            $extSzGroups[$esk] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                        }
                        $extSzGroups[$esk].Add($f)
                    }
                    foreach ($esk in $extSzGroups.Keys) {
                        $esList = $extSzGroups[$esk]
                        if ($esList.Count -lt 2) { continue }
                        foreach ($fi in $esList) {
                            $nf = Get-NormalizedPath -Path $fi.FullName
                            if ($compSha256.ContainsKey($nf)) { continue }
                            $hasCached = $false
                            if ($localDbSnap.ContainsKey($nf)) {
                                $rec = $localDbSnap[$nf]
                                if ($rec["Length"] -eq $fi.Length -and $rec["LastWriteTimeUtcTicks"] -eq $fi.LastWriteTimeUtc.Ticks -and -not [string]::IsNullOrWhiteSpace($rec["Sha256"])) {
                                    $hasCached = $true
                                }
                            }
                            if (-not $hasCached) {
                                $statusSync["Stage"] = "SHA256 (Stage2)"
                                $statusSync["CurrentFile"] = $fi.Name
                                $compSha256[$nf] = (Get-FileHash -LiteralPath $nf -Algorithm SHA256).Hash
                            }
                        }
                    }
                }
            }
        }
        catch {
            $errMsg = $_.Exception.Message
        }

        $sw.Stop()
        [PSCustomObject]@{
            DirectoryPath          = $dirPath
            Files                  = $dirFiles.ToArray()
            ComputedSha256         = $compSha256
            ComputedPayloadSha256  = $compPayload
            ComputedPerceptualHash = $compPerceptual
            ElapsedMs              = $sw.ElapsedMilliseconds
            ErrorMessage           = $errMsg
        }
    }

    $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.Open()
    $activeJobs = [System.Collections.Generic.List[object]]::new()
    $nextDirIdx = 0
    $lastProgressOutputUtc = [datetime]::UtcNow

    try {
        while ($nextDirIdx -lt $totalDirs -or $activeJobs.Count -gt 0) {
            # Fill available slots up to ThrottleLimit
            while ($nextDirIdx -lt $totalDirs -and $activeJobs.Count -lt $ThrottleLimit) {
                $dPath = $scanDirectories[$nextDirIdx]
                $nextDirIdx++
                $dirDbEntries = if ($dbByDir.ContainsKey($dPath)) { $dbByDir[$dPath] } else { @{} }
                $statusSync = [hashtable]::Synchronized(@{
                    Directory   = $dPath
                    Stage       = "Starting"
                    CurrentFile = ""
                    FileTotal   = 0
                    StartedUtc  = [datetime]::UtcNow
                })

                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                [void]$ps.AddScript($workerScript).
                    AddArgument($dPath).
                    AddArgument($dirDbEntries).
                    AddArgument($commonModulePath).
                    AddArgument($excludedSysList).
                    AddArgument($excludedScList).
                    AddArgument($normalizedExtensions).
                    AddArgument($resolvedLogPath).
                    AddArgument($DatabasePath).
                    AddArgument([bool]$NameMatchOnly).
                    AddArgument($statusSync)

                $handle = $ps.BeginInvoke()
                $activeJobs.Add([PSCustomObject]@{
                    PowerShell = $ps
                    Handle     = $handle
                    Status     = $statusSync
                    DirPath    = $dPath
                })
            }

            $anyCompleted = $false
            for ($j = $activeJobs.Count - 1; $j -ge 0; $j--) {
                $job = $activeJobs[$j]
                $elapsedSec = ([datetime]::UtcNow - [datetime]$job.Status["StartedUtc"]).TotalSeconds

                if ($job.Handle.IsCompleted) {
                    $anyCompleted = $true
                    $completedDirs++
                    $lastProgressOutputUtc = [datetime]::UtcNow
                    try {
                        $outCol = $job.PowerShell.EndInvoke($job.Handle)
                        $res = if ($null -ne $outCol -and $outCol.Count -gt 0) { $outCol[0] } else { $null }
                    }
                    catch {
                        $res = [PSCustomObject]@{
                            DirectoryPath = $job.DirPath
                            ErrorMessage  = $_.Exception.Message
                        }
                    }
                    finally {
                        $job.PowerShell.Dispose()
                        $activeJobs.RemoveAt($j)
                    }

                    if ($null -eq $res) { continue }

                    if (-not [string]::IsNullOrWhiteSpace($res.ErrorMessage)) {
                        $scanErrorsCount++
                        Write-SorterLog -Message "Error scanning directory '$($res.DirectoryPath)': $($res.ErrorMessage)" -Level "ERROR" -LogFile $resolvedLogPath
                        continue
                    }

                    $preHashCount = $res.ComputedSha256.Count + $res.ComputedPayloadSha256.Count + $res.ComputedPerceptualHash.Count
                    Write-Host ("[{0}/{1}] Scanned '{2}' ({3} files, {4} pre-hashed, {5} ms)" -f $completedDirs, $totalDirs, $res.DirectoryPath, $res.Files.Count, $preHashCount, $res.ElapsedMs) -ForegroundColor DarkCyan

                    $fileLookup = @{}
                    foreach ($f in $res.Files) {
                        $allScannedFilesList.Add($f)
                        $nf = Get-NormalizedPath -Path $f.FullName
                        $fileLookup[$nf] = $f
                    }
                    foreach ($kv in $res.ComputedSha256.GetEnumerator()) {
                        $hashCache["__precomputed__:$($kv.Key)"] = $kv.Value
                        if (-not $NoDatabase -and $fileLookup.ContainsKey($kv.Key)) {
                            [void](Set-DatabaseFileRecord -Database $sorterDb -FileItem $fileLookup[$kv.Key] -ReferenceRoot $ReferencePath -Sha256 $kv.Value)
                        }
                    }
                    foreach ($kv in $res.ComputedPayloadSha256.GetEnumerator()) {
                        $payloadHashCache["__precomputed__:$($kv.Key)"] = $kv.Value
                        if (-not $NoDatabase -and $fileLookup.ContainsKey($kv.Key)) {
                            [void](Set-DatabaseFileRecord -Database $sorterDb -FileItem $fileLookup[$kv.Key] -ReferenceRoot $ReferencePath -PayloadSha256 $kv.Value)
                        }
                    }
                    foreach ($kv in $res.ComputedPerceptualHash.GetEnumerator()) {
                        $perceptualHashCache["__precomputed__:$($kv.Key)"] = $kv.Value
                        if (-not $NoDatabase -and $fileLookup.ContainsKey($kv.Key)) {
                            [void](Set-DatabaseFileRecord -Database $sorterDb -FileItem $fileLookup[$kv.Key] -ReferenceRoot $ReferencePath -PerceptualHash $kv.Value)
                        }
                    }

                    $unsavedPreHashes += $preHashCount
                    if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath) -and $unsavedPreHashes -ge 50) {
                        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
                        $unsavedPreHashes = 0
                    }
                }
                elseif ($elapsedSec -ge $DirectoryTimeoutSeconds) {
                    $anyCompleted = $true
                    $completedDirs++
                    $scanErrorsCount++
                    $lastProgressOutputUtc = [datetime]::UtcNow
                    $stuckStage = $job.Status["Stage"]
                    $stuckFile = $job.Status["CurrentFile"]
                    $timeoutMsg = "Timeout (${DirectoryTimeoutSeconds}s) scanning directory '$($job.DirPath)' at stage '$stuckStage', file '$stuckFile'. Skipping stuck thread."
                    Write-SorterLog -Message $timeoutMsg -Level "ERROR" -LogFile $resolvedLogPath
                    try { $job.PowerShell.Stop() } catch { }
                    try { $job.PowerShell.Dispose() } catch { }
                    $activeJobs.RemoveAt($j)
                }
            }

            if (-not $anyCompleted) {
                $quietSec = ([datetime]::UtcNow - $lastProgressOutputUtc).TotalSeconds
                if ($quietSec -ge 5 -and $activeJobs.Count -gt 0) {
                    $lastProgressOutputUtc = [datetime]::UtcNow
                    $activeSummaries = @(
                        $activeJobs | Select-Object -First 4 | ForEach-Object {
                            $el = [int]([datetime]::UtcNow - [datetime]$_.Status["StartedUtc"]).TotalSeconds
                            $cf = if (-not [string]::IsNullOrWhiteSpace($_.Status["CurrentFile"])) { ": $($_.Status['CurrentFile'])" } else { "" }
                            "'$($_.DirPath)' [$($_.Status['Stage'])$cf, ${el}s]"
                        }
                    )
                    $moreCount = [Math]::Max(0, $activeJobs.Count - 4)
                    $moreSuffix = if ($moreCount -gt 0) { " (+$moreCount more)" } else { "" }
                    Write-Host ("   [ACTIVE {0}/{1}] Working on {2} thread(s): {3}{4}" -f $completedDirs, $totalDirs, $activeJobs.Count, ($activeSummaries -join "; "), $moreSuffix) -ForegroundColor DarkGray
                }
                Start-Sleep -Milliseconds 50
            }
        }
    }
    finally {
        foreach ($remJob in $activeJobs) {
            try { $remJob.PowerShell.Dispose() } catch { }
        }
        $pool.Close()
        $pool.Dispose()
    }

    if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath) -and $unsavedPreHashes -gt 0) {
        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
    }
}
else {
    foreach ($dirPath in $scanDirectories) {
        try {
            $rawFiles = [System.IO.DirectoryInfo]::new($dirPath).GetFiles()
            foreach ($f in $rawFiles) {
                if ([bool]($f.Attributes -band ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)) -or
                    $f.Name.StartsWith('.')) {
                    continue
                }
                $normFull = Get-NormalizedPath -Path $f.FullName
                $extLow = $f.Extension.ToLowerInvariant()
                if (($normFull -ne $resolvedLogPath) -and
                    ([string]::IsNullOrWhiteSpace($DatabasePath) -or ($normFull -ne $DatabasePath)) -and
                    (-not $excludedSystemFileNames.Contains($f.Name)) -and
                    (-not $excludedSidecarExtensions.Contains($extLow)) -and
                    ($normalizedExtensions.Count -eq 0 -or ($normalizedExtensions -contains $extLow))) {
                    $allScannedFilesList.Add($f)
                }
            }
        }
        catch {
            $scanErrorsCount++
            Write-SorterLog -Message "Error scanning directory '$dirPath': $($_.Exception.Message)" -Level "ERROR" -LogFile $resolvedLogPath
        }
    }
}

$allScannedFiles = @($allScannedFilesList | Sort-Object -Property FullName)

$stats = [ordered]@{
    TotalScanned             = $allScannedFiles.Count
    CandidateNameGroups      = 0
    DuplicateSetsFound       = 0
    DuplicatesMoved          = 0
    HashesComputed           = 0
    PayloadHashesComputed    = 0
    PerceptualHashesComputed = 0
    CacheHits                = 0
    DirectoryHashHits        = 0
    Errors                   = $scanErrorsCount
    LogPath                  = $resolvedLogPath
    DatabasePath             = if ($NoDatabase) { $null } else { $DatabasePath }
}

if ($allScannedFiles.Count -eq 0) {
    Write-SorterLog -Message "No matching files found to scan in '$resolvedRoot'." -Level "WARN" -LogFile $resolvedLogPath
    if ($PassThru) {
        return [PSCustomObject]$stats
    }
    return
}

foreach ($sf in $allScannedFiles) {
    $rawStem = [System.IO.Path]::GetFileNameWithoutExtension($sf.Name)
    $normStem = Get-NormalizedBaseName -BaseName $rawStem
    Write-SorterLog -Message "Scanned file: '$($sf.FullName)' (Ext='$($sf.Extension.ToLowerInvariant())', NormName='$normStem', Size=$($sf.Length) bytes)" -Level "INFO" -LogFile $resolvedLogPath
    if (-not $NoDatabase -and -not $WhatIfPreference) {
        [void](Set-DatabaseFileRecord -Database $sorterDb -FileItem $sf -ReferenceRoot $ReferencePath)
    }
}

# Evaluate directory state hashes against database
$filesByDir = @{}
foreach ($f in $allScannedFiles) {
    $dirNorm = Get-NormalizedPath -Path $f.DirectoryName
    if (-not $filesByDir.ContainsKey($dirNorm)) {
        $filesByDir[$dirNorm] = [System.Collections.Generic.List[object]]::new()
    }
    $filesByDir[$dirNorm].Add($f)
}

$sourceDirsToRefresh = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$destDirsToRefresh = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($dirKey in $filesByDir.Keys) {
    $dirFiles = $filesByDir[$dirKey]
    $computedDirHash = Get-DirectoryStateHash -Files $dirFiles
    if ($sorterDb["SourceDirectories"].ContainsKey($dirKey)) {
        $prevEntry = $sorterDb["SourceDirectories"][$dirKey]
        if ($prevEntry.ContainsKey("DirectoryHash") -and $prevEntry["DirectoryHash"] -eq $computedDirHash) {
            $stats.DirectoryHashHits++
        }
    }
    elseif ($sorterDb["DestinationDirectories"].ContainsKey($dirKey)) {
        $prevEntry = $sorterDb["DestinationDirectories"][$dirKey]
        if ($prevEntry.ContainsKey("DirectoryHash") -and $prevEntry["DirectoryHash"] -eq $computedDirHash) {
            $stats.DirectoryHashHits++
        }
    }
}

$plannedDuplicateTargets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$quarantinedDuplicateFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$keptOriginalFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# ==============================================================================
# STAGE 1: First hint - Group files by (Extension + NormalizedBaseName)
# Different extensions ALWAYS produce distinct keys and are never grouped together.
# ==============================================================================
$nameGroups = @{}

foreach ($file in $allScannedFiles) {
    $extLower = $file.Extension.ToLowerInvariant()
    $rawStem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $normStem = Get-NormalizedBaseName -BaseName $rawStem

    $groupKey = "{0}|{1}" -f $extLower, $normStem
    if (-not $nameGroups.ContainsKey($groupKey)) {
        $nameGroups[$groupKey] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    }
    $nameGroups[$groupKey].Add($file)
}

foreach ($groupKey in $nameGroups.Keys) {
    $candidates = $nameGroups[$groupKey]
    if ($candidates.Count -lt 2) {
        continue
    }

    $stats.CandidateNameGroups++
    $parts = $groupKey -split '\|', 2
    $extForGroup = $parts[0]
    $normStemForGroup = if ($parts.Count -gt 1) { $parts[1] } else { "" }

    $candidateNamesLog = ($candidates | ForEach-Object { "'$($_.FullName)' ($($_.Length) B)" }) -join ", "
    Write-SorterLog -Message "Stage 1 Candidate Name Group [$groupKey]: $candidateNamesLog" -Level "INFO" -LogFile $resolvedLogPath

    # --- STEP 1A: Compare candidates with identical byte size using whole-file SHA-256 ---
    $sizeGroups = @{}
    foreach ($cand in $candidates) {
        $lenKey = [string]$cand.Length
        if (-not $sizeGroups.ContainsKey($lenKey)) {
            $sizeGroups[$lenKey] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        }
        $sizeGroups[$lenKey].Add($cand)
    }

    foreach ($lenKey in $sizeGroups.Keys) {
        $sameSizeFiles = $sizeGroups[$lenKey]
        if ($sameSizeFiles.Count -lt 2) {
            continue
        }

        $hashGroups = @{}
        foreach ($fileItem in $sameSizeFiles) {
            try {
                $normFilePath = Get-NormalizedPath -Path $fileItem.FullName
                $cacheHitsBefore = $stats.CacheHits
                $wasInSessionCache = $hashCache.ContainsKey($normFilePath)

                $fileSha = Get-CachedFileHash `
                    -FilePath $normFilePath `
                    -Cache $hashCache `
                    -Database $sorterDb `
                    -Stats $stats `
                    -FileItem $fileItem `
                    -ReferenceRoot $ReferencePath

                if (-not $wasInSessionCache -and ($stats.CacheHits -eq $cacheHitsBefore)) {
                    $stats.HashesComputed++
                    $totalNewHashes = $stats.HashesComputed + $stats.PayloadHashesComputed
                    if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath) -and ($totalNewHashes % 250 -eq 0)) {
                        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
                    }
                }

                if (-not $hashGroups.ContainsKey($fileSha)) {
                    $hashGroups[$fileSha] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                }
                $hashGroups[$fileSha].Add($fileItem)
            }
            catch {
                $stats.Errors++
                Write-SorterLog -Message "Failed to compute hash for '$($fileItem.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $resolvedLogPath
            }
        }

        foreach ($shaKey in $hashGroups.Keys) {
            $identicalFiles = $hashGroups[$shaKey]
            if ($identicalFiles.Count -lt 2) {
                continue
            }

            Move-ConfirmedDuplicateGroup `
                -IdenticalFiles $identicalFiles `
                -Sha256Hash $shaKey `
                -DetectionStage "Stage1:Name+Content" `
                -PreferredStem $normStemForGroup `
                -DupFolderName $DuplicateFolderName `
                -PlannedTargets $plannedDuplicateTargets `
                -QuarantinedFiles $quarantinedDuplicateFiles `
                -KeptOriginals $keptOriginalFiles `
                -Database $sorterDb `
                -SourceDirsToRefresh $sourceDirsToRefresh `
                -DestDirsToRefresh $destDirsToRefresh `
                -Stats $stats `
                -LogFile $resolvedLogPath `
                -WhatIf:$WhatIfPreference
        }
    }

    # --- STEP 1B: Compare remaining candidates in the same name group using metadata-stripped Image Payload SHA-256 ---
    # This catches duplicate images that differ in byte size by a few KB due to edited EXIF/XMP headers
    # (e.g. 2025-01-02_153731.jpg vs 2025-01-02_153731_000.jpg or 2025-01-02_153726_000.jpg vs 2025-01-02_153726-01.jpg)
    # while keeping genuine burst shots intact.
    if ($extForGroup -in @(".jpg", ".jpeg", ".webp", ".png", ".dng", ".tif", ".tiff", ".cr2", ".nef", ".arw", ".orf", ".pef")) {
        $remainingInNameGroup = @(
            $candidates | Where-Object {
                -not $quarantinedDuplicateFiles.Contains((Get-NormalizedPath -Path $_.FullName))
            }
        )

        if ($remainingInNameGroup.Count -ge 2) {
            $payloadGroups = @{}
            foreach ($fileItem in $remainingInNameGroup) {
                try {
                    $normFilePath = Get-NormalizedPath -Path $fileItem.FullName
                    $cacheHitsBefore = $stats.CacheHits
                    $wasInPayloadSessionCache = $payloadHashCache.ContainsKey($normFilePath)

                    $payloadSha = Get-CachedImagePayloadHash `
                        -FilePath $normFilePath `
                        -PayloadCache $payloadHashCache `
                        -Database $sorterDb `
                        -Stats $stats `
                        -FileItem $fileItem `
                        -ReferenceRoot $ReferencePath

                    if ([string]::IsNullOrWhiteSpace($payloadSha)) {
                        continue
                    }

                    if (-not $wasInPayloadSessionCache -and ($stats.CacheHits -eq $cacheHitsBefore)) {
                        $stats.PayloadHashesComputed++
                        $totalNewHashes = $stats.HashesComputed + $stats.PayloadHashesComputed
                        if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath) -and ($totalNewHashes % 250 -eq 0)) {
                            Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
                        }
                    }

                    if (-not $payloadGroups.ContainsKey($payloadSha)) {
                        $payloadGroups[$payloadSha] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                    }
                    $payloadGroups[$payloadSha].Add($fileItem)
                }
                catch {
                    $stats.Errors++
                    Write-SorterLog -Message "Failed to compute image payload hash for '$($fileItem.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $resolvedLogPath
                }
            }

            foreach ($pShaKey in $payloadGroups.Keys) {
                $samePayloadFiles = $payloadGroups[$pShaKey]
                if ($samePayloadFiles.Count -lt 2) {
                    continue
                }

                Move-ConfirmedDuplicateGroup `
                    -IdenticalFiles $samePayloadFiles `
                    -Sha256Hash $pShaKey `
                    -DetectionStage "Stage1:Name+ImagePayload" `
                    -PreferredStem $normStemForGroup `
                    -DupFolderName $DuplicateFolderName `
                    -PlannedTargets $plannedDuplicateTargets `
                    -QuarantinedFiles $quarantinedDuplicateFiles `
                    -KeptOriginals $keptOriginalFiles `
                    -Database $sorterDb `
                    -SourceDirsToRefresh $sourceDirsToRefresh `
                    -DestDirsToRefresh $destDirsToRefresh `
                    -Stats $stats `
                    -LogFile $resolvedLogPath `
                    -WhatIf:$WhatIfPreference
            }
        }

        # --- STEP 1C: Compare remaining candidates in the same name group using Perceptual Hash (dHash 64-bit) ---
        # This catches duplicate photos saved at different resolutions or JPEG/DNG compression levels
        # (e.g. 2025-01-02_160119_001.jpg [18 KB thumbnail] vs 2025-01-02_160119_003.jpg [67 KB full image]).
        $remainingForPerceptual = @(
            $candidates | Where-Object {
                -not $quarantinedDuplicateFiles.Contains((Get-NormalizedPath -Path $_.FullName))
            }
        )

        if ($remainingForPerceptual.Count -ge 2) {
            $perceptualCandidates = [System.Collections.Generic.List[object]]::new()

            foreach ($fileItem in $remainingForPerceptual) {
                try {
                    $normFilePath = Get-NormalizedPath -Path $fileItem.FullName
                    $cacheHitsBefore = $stats.CacheHits
                    $wasInPerceptualSessionCache = $perceptualHashCache.ContainsKey($normFilePath)

                    $pHash = Get-CachedImagePerceptualHash `
                        -FilePath $normFilePath `
                        -PerceptualCache $perceptualHashCache `
                        -Database $sorterDb `
                        -Stats $stats `
                        -FileItem $fileItem `
                        -ReferenceRoot $ReferencePath

                    if ([string]::IsNullOrWhiteSpace($pHash)) {
                        continue
                    }

                    if (-not $wasInPerceptualSessionCache -and ($stats.CacheHits -eq $cacheHitsBefore)) {
                        $stats.PerceptualHashesComputed++
                        $totalNewHashes = $stats.HashesComputed + $stats.PayloadHashesComputed + $stats.PerceptualHashesComputed
                        if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath) -and ($totalNewHashes % 250 -eq 0)) {
                            Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
                        }
                    }

                    $perceptualCandidates.Add([PSCustomObject]@{
                        File           = $fileItem
                        PerceptualHash = $pHash
                    })
                }
                catch {
                    $stats.Errors++
                    Write-SorterLog -Message "Failed to compute perceptual hash for '$($fileItem.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $resolvedLogPath
                }
            }

            if ($perceptualCandidates.Count -ge 2) {
                # Sort by file Length descending so the highest-quality file anchors each visual cluster
                $sortedPerceptual = @($perceptualCandidates | Sort-Object -Property @{ Expression = { $_.File.Length }; Ascending = $false })
                $clusteredIdx = [System.Collections.Generic.HashSet[int]]::new()

                for ($i = 0; $i -lt $sortedPerceptual.Count; $i++) {
                    if ($clusteredIdx.Contains($i)) {
                        continue
                    }

                    $anchor = $sortedPerceptual[$i]
                    $visualGroup = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                    $visualGroup.Add($anchor.File)
                    [void]$clusteredIdx.Add($i)

                    for ($j = $i + 1; $j -lt $sortedPerceptual.Count; $j++) {
                        if ($clusteredIdx.Contains($j)) {
                            continue
                        }

                        $other = $sortedPerceptual[$j]
                        $hammingDist = Get-PerceptualHashDistance -HashA $anchor.PerceptualHash -HashB $other.PerceptualHash
                        if ($hammingDist -le $MaxPerceptualDistance) {
                            $visualGroup.Add($other.File)
                            [void]$clusteredIdx.Add($j)
                        }
                    }

                    if ($visualGroup.Count -ge 2) {
                        Move-ConfirmedDuplicateGroup `
                            -IdenticalFiles $visualGroup `
                            -Sha256Hash $anchor.PerceptualHash `
                            -DetectionStage "Stage1:Name+PerceptualHash" `
                            -PreferredStem $normStemForGroup `
                            -DupFolderName $DuplicateFolderName `
                            -PlannedTargets $plannedDuplicateTargets `
                            -QuarantinedFiles $quarantinedDuplicateFiles `
                            -KeptOriginals $keptOriginalFiles `
                            -Database $sorterDb `
                            -SourceDirsToRefresh $sourceDirsToRefresh `
                            -DestDirsToRefresh $destDirsToRefresh `
                            -Stats $stats `
                            -LogFile $resolvedLogPath `
                            -WhatIf:$WhatIfPreference
                    }
                }
            }
        }
    }
}

# ==============================================================================
# STAGE 2: Second check - Content verification (Extension + Size + SHA-256)
# for remaining files within Path that had different filenames in Stage 1.
# ==============================================================================
if (-not $NameMatchOnly) {
    $extSizeGroups = @{}

    foreach ($file in $allScannedFiles) {
        $normPath = Get-NormalizedPath -Path $file.FullName
        if ($quarantinedDuplicateFiles.Contains($normPath)) {
            continue
        }

        $extLower = $file.Extension.ToLowerInvariant()
        $extSizeKey = "{0}|{1}" -f $extLower, $file.Length
        if (-not $extSizeGroups.ContainsKey($extSizeKey)) {
            $extSizeGroups[$extSizeKey] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        }
        $extSizeGroups[$extSizeKey].Add($file)
    }

    foreach ($extSizeKey in $extSizeGroups.Keys) {
        $sameExtAndSizeFiles = $extSizeGroups[$extSizeKey]
        if ($sameExtAndSizeFiles.Count -lt 2) {
            continue
        }

        $candidateLog = ($sameExtAndSizeFiles | ForEach-Object { "'$($_.FullName)'" }) -join ", "
        Write-SorterLog -Message "Stage 2 Content Candidate Group (Same Extension & Size [$extSizeKey]): $candidateLog" -Level "INFO" -LogFile $resolvedLogPath

        $hashGroups = @{}
        foreach ($fileItem in $sameExtAndSizeFiles) {
            try {
                $normFilePath = Get-NormalizedPath -Path $fileItem.FullName
                $cacheHitsBefore = $stats.CacheHits
                $wasInSessionCache = $hashCache.ContainsKey($normFilePath)

                $fileSha = Get-CachedFileHash `
                    -FilePath $normFilePath `
                    -Cache $hashCache `
                    -Database $sorterDb `
                    -Stats $stats `
                    -FileItem $fileItem `
                    -ReferenceRoot $ReferencePath

                if (-not $wasInSessionCache -and ($stats.CacheHits -eq $cacheHitsBefore)) {
                    $stats.HashesComputed++
                    $totalNewHashes = $stats.HashesComputed + $stats.PayloadHashesComputed
                    if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath) -and ($totalNewHashes % 250 -eq 0)) {
                        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
                    }
                }

                if (-not $hashGroups.ContainsKey($fileSha)) {
                    $hashGroups[$fileSha] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
                }
                $hashGroups[$fileSha].Add($fileItem)
            }
            catch {
                $stats.Errors++
                Write-SorterLog -Message "Failed to compute hash for '$($fileItem.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $resolvedLogPath
            }
        }

        foreach ($shaKey in $hashGroups.Keys) {
            $identicalFiles = $hashGroups[$shaKey]
            if ($identicalFiles.Count -lt 2) {
                continue
            }

            Move-ConfirmedDuplicateGroup `
                -IdenticalFiles $identicalFiles `
                -Sha256Hash $shaKey `
                -DetectionStage "Stage2:Content" `
                -PreferredStem "" `
                -DupFolderName $DuplicateFolderName `
                -PlannedTargets $plannedDuplicateTargets `
                -QuarantinedFiles $quarantinedDuplicateFiles `
                -KeptOriginals $keptOriginalFiles `
                -Database $sorterDb `
                -SourceDirsToRefresh $sourceDirsToRefresh `
                -DestDirsToRefresh $destDirsToRefresh `
                -Stats $stats `
                -LogFile $resolvedLogPath `
                -WhatIf:$WhatIfPreference
        }
    }
}

# ==============================================================================
# STAGE 3: Reference Library / Database check (e.g. against X:\Pictures)
# Detects files in Path whose original copy was already moved to ReferencePath
# and recorded in DestinationFiles of .smartmediasorter_db.json.
# ==============================================================================
if (-not $NoDatabase -and $sorterDb["DestinationFiles"].Count -gt 0) {
    # Index external DestinationFiles by "extLower|Length" for O(1) candidate lookup
    $destByExtAndLength = @{}
    foreach ($destKey in $sorterDb["DestinationFiles"].Keys) {
        $normDestKey = Get-NormalizedPath -Path $destKey
        # Skip files that are inside the currently scanned root (already compared in Stage 1 & 2)
        if ($normDestKey.StartsWith($resolvedRoot + "\", [System.StringComparison]::OrdinalIgnoreCase) -or
            $normDestKey.Equals($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $destRec = $sorterDb["DestinationFiles"][$destKey]
        $destExt = [System.IO.Path]::GetExtension($normDestKey).ToLowerInvariant()
        $destLen = if ($destRec.ContainsKey("Length")) { [long]$destRec["Length"] } else { -1L }
        if ($destLen -lt 0) {
            continue
        }

        $lookupKey = "{0}|{1}" -f $destExt, $destLen
        if (-not $destByExtAndLength.ContainsKey($lookupKey)) {
            $destByExtAndLength[$lookupKey] = [System.Collections.Generic.List[string]]::new()
        }
        $destByExtAndLength[$lookupKey].Add($normDestKey)
    }

    foreach ($file in $allScannedFiles) {
        $normFilePath = Get-NormalizedPath -Path $file.FullName
        if ($quarantinedDuplicateFiles.Contains($normFilePath)) {
            continue
        }

        $extLower = $file.Extension.ToLowerInvariant()
        $lookupKey = "{0}|{1}" -f $extLower, $file.Length
        if (-not $destByExtAndLength.ContainsKey($lookupKey)) {
            continue
        }

        $matchingDestPaths = $destByExtAndLength[$lookupKey]
        $matchedDestOriginal = $null
        $fileSha = $null

        foreach ($candDestPath in $matchingDestPaths) {
            if (-not (Test-Path -LiteralPath $candDestPath -PathType Leaf)) {
                continue
            }

            if ($null -eq $fileSha) {
                $cacheHitsBefore = $stats.CacheHits
                $wasInSessionCache = $hashCache.ContainsKey($normFilePath)
                $fileSha = Get-CachedFileHash `
                    -FilePath $normFilePath `
                    -Cache $hashCache `
                    -Database $sorterDb `
                    -Stats $stats `
                    -FileItem $file `
                    -ReferenceRoot $ReferencePath
                if (-not $wasInSessionCache -and ($stats.CacheHits -eq $cacheHitsBefore)) {
                    $stats.HashesComputed++
                }
            }

            $cacheHitsBeforeDest = $stats.CacheHits
            $wasDestInSessionCache = $hashCache.ContainsKey($candDestPath)
            $destSha = Get-CachedFileHash `
                -FilePath $candDestPath `
                -Cache $hashCache `
                -Database $sorterDb `
                -Stats $stats `
                -ReferenceRoot $ReferencePath
            if (-not $wasDestInSessionCache -and ($stats.CacheHits -eq $cacheHitsBeforeDest)) {
                $stats.HashesComputed++
            }

            if ($fileSha -eq $destSha) {
                $matchedDestOriginal = $candDestPath
                break
            }
        }

        if ($null -ne $matchedDestOriginal) {
            $stats.DuplicateSetsFound++
            $origName = [System.IO.Path]::GetFileName($matchedDestOriginal)
            Move-SingleDuplicateToLocalSubfolder `
                -DuplicateFile $file `
                -OriginalFileFullPath $matchedDestOriginal `
                -OriginalFileName $origName `
                -Sha256Hash $fileSha `
                -DetectionStage "Stage3:ReferenceDb" `
                -DupFolderName $DuplicateFolderName `
                -PlannedTargets $plannedDuplicateTargets `
                -QuarantinedFiles $quarantinedDuplicateFiles `
                -Database $sorterDb `
                -SourceDirsToRefresh $sourceDirsToRefresh `
                -DestDirsToRefresh $destDirsToRefresh `
                -Stats $stats `
                -LogFile $resolvedLogPath `
                -WhatIf:$WhatIfPreference
        }
    }
}

# Persist updated directory and file hashes to database
if (-not $NoDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath)) {
    if (-not $WhatIfPreference) {
        $normRefRoot = if (-not [string]::IsNullOrWhiteSpace($ReferencePath)) {
            Get-NormalizedPath -Path $ReferencePath
        } else {
            ""
        }

        foreach ($dirKey in $filesByDir.Keys) {
            $isDestDir = $sorterDb["DestinationDirectories"].ContainsKey($dirKey)
            if (-not $isDestDir -and -not [string]::IsNullOrWhiteSpace($normRefRoot)) {
                $isDestDir = (
                    $dirKey.Equals($normRefRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
                    $dirKey.StartsWith($normRefRoot + "\", [System.StringComparison]::OrdinalIgnoreCase)
                )
            }

            if ($isDestDir) {
                [void]$destDirsToRefresh.Add($dirKey)
            }
            else {
                [void]$sourceDirsToRefresh.Add($dirKey)
            }
        }

        Update-DirectoryHashesOnDisk `
            -Database $sorterDb `
            -SourceDirsToRefresh $sourceDirsToRefresh `
            -DestDirsToRefresh $destDirsToRefresh `
            -NormalizedExtensions $normalizedExtensions

        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
        Write-SorterLog -Message "Updated shared hash database at '$DatabasePath'." -Level "INFO" -LogFile $resolvedLogPath
    }
    elseif (($stats.HashesComputed + $stats.PayloadHashesComputed + $stats.PerceptualHashesComputed) -gt 0) {
        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $resolvedLogPath
        Write-SorterLog -Message "Saved newly computed file hashes ($($stats.HashesComputed) SHA-256, $($stats.PayloadHashesComputed) PayloadSha256, $($stats.PerceptualHashesComputed) PerceptualHash) to database '$DatabasePath' during WhatIf run." -Level "INFO" -LogFile $resolvedLogPath
    }
}

$summaryLine = "Completed Find-MediaDuplicates: Scanned=$($stats.TotalScanned), CandidateNameGroups=$($stats.CandidateNameGroups), DuplicateSetsFound=$($stats.DuplicateSetsFound), DuplicatesMoved=$($stats.DuplicatesMoved), HashesComputed=$($stats.HashesComputed), PayloadHashesComputed=$($stats.PayloadHashesComputed), PerceptualHashesComputed=$($stats.PerceptualHashesComputed), CacheHits=$($stats.CacheHits), DirHashHits=$($stats.DirectoryHashHits), Errors=$($stats.Errors). Log='$($stats.LogPath)'"
Write-SorterLog -Message $summaryLine -Level "INFO" -LogFile $resolvedLogPath

Write-Host ""
Write-Host "=== Find-MediaDuplicates Summary ===" -ForegroundColor Cyan
Write-Host "Total scanned         : $($stats.TotalScanned)"
Write-Host "Candidate name groups : $($stats.CandidateNameGroups)"
Write-Host "Duplicate sets found  : $($stats.DuplicateSetsFound)" -ForegroundColor Yellow
Write-Host "Duplicates moved      : $($stats.DuplicatesMoved)" -ForegroundColor Green
Write-Host "SHA-256 computed      : $($stats.HashesComputed) (Payload: $($stats.PayloadHashesComputed), Perceptual: $($stats.PerceptualHashesComputed), Cache hits: $($stats.CacheHits), DirHash hits: $($stats.DirectoryHashHits))" -ForegroundColor Gray
Write-Host "Errors                : $($stats.Errors)" -ForegroundColor $(if ($stats.Errors -gt 0) { "Red" } else { "Gray" })
Write-Host "Database file         : $(if ($NoDatabase) { '(disabled)' } else { $stats.DatabasePath })"
Write-Host "Log file              : $($stats.LogPath)"
Write-Host "====================================" -ForegroundColor Cyan

if ($PassThru) {
    return [PSCustomObject]$stats
}

#endregion Main Execution
