#Requires -Version 7.0

<#
.SYNOPSIS
    Moves or copies photos from a source directory to a structured destination directory
    based on EXIF date metadata, renaming files, detecting duplicates, and maintaining
    a persistent directory/file hash database to accelerate subsequent runs.

.DESCRIPTION
    SmartMediaSorter scans a source directory (default: X:\Photos) for image files,
    extracts EXIF creation timestamps using ExifTool (with automatic fallback to
    native .NET EXIF header parsing and finally file LastWriteTime), and moves
    (or copies) them into a destination directory (default: X:\Pictures) organized as:
        \YYYY\YYYY-MM\YYYY-MM-DD_HHMMSS_XXX.<ext>

    Duplicate detection:
    When one or more files with the same timestamp (YYYY-MM-DD_HHMMSS) and extension
    already exist in the target month folder, the script compares file sizes and
    SHA-256 hashes:
    - If an identical SHA-256 hash is found, the file is skipped and left in the
      source directory, and the event is recorded in the log.
    - If the content differs (e.g., burst shots taken within the same second),
      the next available 3-digit sequence number (_000, _001, _002, ...) is assigned.

    Database acceleration:
    By default, SmartMediaSorter maintains a JSON database (.smartmediasorter_db.json
    in DestinationPath) storing SHA-256 state hashes of source and destination
    directories along with cached per-file EXIF dates and SHA-256 content hashes.
    On subsequent runs, unchanged directories and files reuse cached metadata and hashes,
    and moving/copying files updates the source and destination hashes automatically.

.PARAMETER SourcePath
    The root directory containing photos to process. Defaults to "X:\Photos".

.PARAMETER DestinationPath
    The root directory where organized photos will be stored. Defaults to "X:\Pictures".

.PARAMETER CopyOnly
    If specified, files are copied instead of moved. By default, files are moved.

.PARAMETER Recurse
    Whether to scan subdirectories within SourcePath. Defaults to $true.

.PARAMETER ExifToolPath
    Path or command name for the ExifTool executable. Defaults to "exiftool".

.PARAMETER LogPath
    Optional path to the log file. Defaults to "SmartMediaSorter_<timestamp>.log"
    in the script directory.

.PARAMETER DatabasePath
    Optional path to the persistent JSON database file. Defaults to
    "<DestinationPath>\.smartmediasorter_db.json".

.PARAMETER NoDatabase
    If specified, disables reading from and writing to the persistent hash database.

.PARAMETER RebuildDatabase
    If specified, ignores any existing database file and rebuilds all directory and
    file hashes from scratch.

.PARAMETER IncludeExtensions
    List of file extensions (with leading dot) to process.

.PARAMETER PassThru
    If specified, returns a summary object with processing statistics at the end.

.EXAMPLE
    .\SmartMediaSorter.ps1 -WhatIf
    Simulates moving photos from X:\Photos to X:\Pictures without modifying any files.

.EXAMPLE
    .\SmartMediaSorter.ps1 -SourcePath "X:\Photos" -DestinationPath "X:\Pictures"
    Moves photos from X:\Photos to X:\Pictures, renaming, skipping duplicates, and updating the database.

.EXAMPLE
    .\SmartMediaSorter.ps1 -RebuildDatabase
    Forces a full rescan and rebuilds the source and destination hash database.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$SourcePath = "X:\Photos",

    [Parameter(Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath = "X:\Pictures",

    [Parameter()]
    [switch]$CopyOnly,

    [Parameter()]
    [bool]$Recurse = $true,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExifToolPath = "exiftool",

    [Parameter()]
    [string]$LogPath = "",

    [Parameter()]
    [string]$DatabasePath = "",

    [Parameter()]
    [switch]$NoDatabase,

    [Parameter()]
    [switch]$RebuildDatabase,

    [Parameter()]
    [string[]]$IncludeExtensions = @(
        ".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp",
        ".tif", ".tiff", ".bmp", ".gif", ".dng", ".cr2",
        ".cr3", ".nef", ".arw", ".orf", ".rw2", ".pef", ".srw"
    ),

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

try {
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
}
catch {
    # System.Drawing is optional if ExifTool handles all files
}

#region Helper Functions

function Test-ExifToolExecutable {
    <#
    .SYNOPSIS
        Verifies that an ExifTool binary path exists and can execute without OS loader errors.
    #>
    param([Parameter(Mandatory = $true)][string]$BinaryPath)

    $prevEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $verOutput = & $BinaryPath -ver 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace(($verOutput -join ""))) {
            return $true
        }
    }
    catch {
        return $false
    }
    finally {
        $ErrorActionPreference = $prevEap
    }
    return $false
}

function Resolve-ExifToolBinary {
    <#
    .SYNOPSIS
        Locates a working ExifTool executable from parameter, PATH, or common install paths.
        Returns $null if no working binary is available (allowing native .NET EXIF fallback).
    #>
    param([string]$ConfiguredPath)

    $candidates = [System.Collections.Generic.List[string]]::new()

    $cmd = Get-Command -Name $ConfiguredPath -ErrorAction SilentlyContinue
    if ($null -ne $cmd -and $cmd.Source) {
        $candidates.Add($cmd.Source)
    }

    if (Test-Path -LiteralPath $ConfiguredPath -PathType Leaf) {
        $candidates.Add((Resolve-Path -LiteralPath $ConfiguredPath).Path)
    }

    $fallbackPaths = @(
        (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\WinGet\Links\exiftool.exe"),
        (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Programs\ExifToolGUI\exiftool.exe"),
        "C:\Windows\exiftool.exe"
    )

    foreach ($fp in $fallbackPaths) {
        if ($fp -and (Test-Path -LiteralPath $fp -PathType Leaf)) {
            $candidates.Add($fp)
        }
    }

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (Test-ExifToolExecutable -BinaryPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function ConvertFrom-ExifDateString {
    <#
    .SYNOPSIS
        Parses an EXIF date string into a [datetime] object, or returns $null if invalid.
    #>
    param([AllowNull()][object]$RawValue)

    if ($null -eq $RawValue) {
        return $null
    }

    $str = (([string]$RawValue) -replace '\x00', '').Trim()
    if ([string]::IsNullOrWhiteSpace($str) -or $str.StartsWith("0000")) {
        return $null
    }

    # Strip optional subseconds or timezone suffixes if present
    if ($str -match '^(\d{4}[-:]\d{2}[-:]\d{2}\s+\d{2}:\d{2}:\d{2})') {
        $str = $Matches[1] -replace '^(\d{4}):(\d{2}):(\d{2})', '$1-$2-$3'
    }

    $parsedDate = [datetime]::MinValue
    [string[]]$formats = @(
        "yyyy-MM-dd HH:mm:ss",
        "yyyy:MM:dd HH:mm:ss"
    )

    if ([datetime]::TryParseExact(
            $str,
            $formats,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$parsedDate
        )) {
        if ($parsedDate.Year -ge 1900 -and $parsedDate.Year -le 2100) {
            return $parsedDate
        }
    }

    return $null
}

function Read-UInt16Endian {
    param([byte[]]$Buffer, [int]$Offset, [bool]$LittleEndian)
    if ($LittleEndian) {
        return [int][System.BitConverter]::ToUInt16($Buffer, $Offset)
    }
    return ([int]$Buffer[$Offset] -shl 8) -bor [int]$Buffer[$Offset + 1]
}

function Read-UInt32Endian {
    param([byte[]]$Buffer, [int]$Offset, [bool]$LittleEndian)
    if ($LittleEndian) {
        return [int][System.BitConverter]::ToUInt32($Buffer, $Offset)
    }
    return ([int]$Buffer[$Offset] -shl 24) -bor
           ([int]$Buffer[$Offset + 1] -shl 16) -bor
           ([int]$Buffer[$Offset + 2] -shl 8) -bor
           [int]$Buffer[$Offset + 3]
}

function Get-DotNetExifDate {
    <#
    .SYNOPSIS
        Reads EXIF DateTimeOriginal (0x9003) or CreateDate (0x9004) directly from
        JPEG APP1 / TIFF EXIF headers without decoding pixel data.
    #>
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $fs = $null
    try {
        $fs = [System.IO.File]::Open(
            $FilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )

        $maxRead = [Math]::Min([int]$fs.Length, 131072)
        if ($maxRead -lt 16) {
            return $null
        }

        $buf = [byte[]]::new($maxRead)
        $bytesRead = $fs.Read($buf, 0, $maxRead)

        # Check for JPEG SOI (0xFF 0xD8)
        if ($buf[0] -eq 0xFF -and $buf[1] -eq 0xD8) {
            $pos = 2
            while ($pos + 4 -le $bytesRead) {
                if ($buf[$pos] -ne 0xFF) {
                    break
                }
                $marker = $buf[$pos + 1]
                # Skip padding 0xFF bytes
                if ($marker -eq 0xFF) {
                    $pos++
                    continue
                }
                # SOS (0xDA) or EOI (0xD9) means end of metadata segments
                if ($marker -eq 0xDA -or $marker -eq 0xD9) {
                    break
                }

                $segLen = ([int]$buf[$pos + 2] -shl 8) -bor [int]$buf[$pos + 3]
                if ($segLen -lt 2) {
                    break
                }

                # Check APP1 (0xE1) with "Exif\0\0" signature
                if ($marker -eq 0xE1 -and ($pos + 10 -le $bytesRead) -and
                    $buf[$pos + 4] -eq 0x45 -and $buf[$pos + 5] -eq 0x78 -and
                    $buf[$pos + 6] -eq 0x69 -and $buf[$pos + 7] -eq 0x66 -and
                    $buf[$pos + 8] -eq 0x00 -and $buf[$pos + 9] -eq 0x00) {

                    $tiffStart = $pos + 10
                    if ($tiffStart + 8 -le $bytesRead) {
                        $isLittleEndian = ($buf[$tiffStart] -eq 0x49 -and $buf[$tiffStart + 1] -eq 0x49)
                        $ifd0RelOffset = Read-UInt32Endian -Buffer $buf -Offset ($tiffStart + 4) -LittleEndian $isLittleEndian

                        $ifdOffsetsToScan = [System.Collections.Generic.Queue[int]]::new()
                        $ifdOffsetsToScan.Enqueue($ifd0RelOffset)

                        $foundDates = @{}

                        while ($ifdOffsetsToScan.Count -gt 0) {
                            $relIfd = $ifdOffsetsToScan.Dequeue()
                            $absIfd = $tiffStart + $relIfd
                            if ($absIfd + 2 -gt $bytesRead) {
                                continue
                            }

                            $entryCount = Read-UInt16Endian -Buffer $buf -Offset $absIfd -LittleEndian $isLittleEndian
                            for ($i = 0; $i -lt $entryCount; $i++) {
                                $entryPos = $absIfd + 2 + ($i * 12)
                                if ($entryPos + 12 -gt $bytesRead) {
                                    break
                                }

                                $tagId = Read-UInt16Endian -Buffer $buf -Offset $entryPos -LittleEndian $isLittleEndian
                                $fieldType = Read-UInt16Endian -Buffer $buf -Offset ($entryPos + 2) -LittleEndian $isLittleEndian
                                $valCount = Read-UInt32Endian -Buffer $buf -Offset ($entryPos + 4) -LittleEndian $isLittleEndian
                                $valOffset = Read-UInt32Endian -Buffer $buf -Offset ($entryPos + 8) -LittleEndian $isLittleEndian

                                # 34665 (0x08769) = ExifSubIFD pointer
                                if ($tagId -eq 0x08769 -and $valOffset -gt 0) {
                                    $ifdOffsetsToScan.Enqueue($valOffset)
                                }
                                # 36867 (0x09003) = DateTimeOriginal, 36868 (0x09004) = CreateDate, 306 (0x00132) = ModifyDate (ASCII type = 2)
                                elseif (($tagId -eq 0x09003 -or $tagId -eq 0x09004 -or $tagId -eq 0x00132) -and $fieldType -eq 2 -and $valCount -ge 19) {
                                    $strAbsOffset = $tiffStart + $valOffset
                                    if ($strAbsOffset + 19 -le $bytesRead) {
                                        $rawAscii = [System.Text.Encoding]::ASCII.GetString($buf, $strAbsOffset, [Math]::Min($valCount, 20))
                                        $parsed = ConvertFrom-ExifDateString -RawValue $rawAscii
                                        if ($null -ne $parsed) {
                                            $foundDates[[int]$tagId] = $parsed
                                        }
                                    }
                                }
                            }
                        }

                        if ($foundDates.ContainsKey(0x09003)) {
                            return [PSCustomObject]@{ Date = $foundDates[0x09003]; Source = "EXIF:DateTimeOriginal" }
                        }
                        if ($foundDates.ContainsKey(0x09004)) {
                            return [PSCustomObject]@{ Date = $foundDates[0x09004]; Source = "EXIF:CreateDate" }
                        }
                        if ($foundDates.ContainsKey(0x00132)) {
                            return [PSCustomObject]@{ Date = $foundDates[0x00132]; Source = "EXIF:ModifyDate" }
                        }
                    }
                }

                $pos += 2 + $segLen
            }
        }
    }
    catch {
        return $null
    }
    finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }

    return $null
}

function Get-BatchExifMetadata {
    <#
    .SYNOPSIS
        Executes ExifTool once for the specified target paths (directories or files)
        and returns a hashtable mapping normalized file paths to EXIF metadata objects.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$BinaryPath,
        [Parameter(Mandatory = $true)][string[]]$TargetPaths,
        [Parameter()][bool]$Recursive = $false
    )

    $metadataMap = @{}
    if ($null -eq $TargetPaths -or $TargetPaths.Count -eq 0) {
        return $metadataMap
    }

    $argsList = @(
        "-q",
        "-q",
        "-j",
        "-d", "%Y-%m-%d %H:%M:%S",
        "-DateTimeOriginal",
        "-CreateDate",
        "-MediaCreateDate",
        "-TrackCreateDate"
    )

    if ($Recursive) {
        $argsList += "-r"
    }

    $argsList += $TargetPaths

    $previousEncoding = [Console]::OutputEncoding
    $prevEap = $ErrorActionPreference
    $rawOutput = $null
    try {
        $ErrorActionPreference = "Continue"
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $rawOutput = & $BinaryPath @argsList 2>$null
    }
    catch {
        return $metadataMap
    }
    finally {
        [Console]::OutputEncoding = $previousEncoding
        $ErrorActionPreference = $prevEap
    }

    if ($null -eq $rawOutput) {
        return $metadataMap
    }

    $jsonText = ($rawOutput -join [Environment]::NewLine).Trim()
    if ([string]::IsNullOrWhiteSpace($jsonText) -or -not $jsonText.StartsWith("[")) {
        return $metadataMap
    }

    $parsedJson = ConvertFrom-Json -InputObject $jsonText
    foreach ($item in $parsedJson) {
        if ($null -ne $item -and $null -ne $item.PSObject.Properties["SourceFile"] -and $item.SourceFile) {
            $normPath = Get-NormalizedPath -Path ([string]$item.SourceFile)
            $metadataMap[$normPath] = $item
        }
    }

    return $metadataMap
}

#endregion Helper Functions

#region Main Execution

# Initialize log file path
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $runTimestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $LogPath = Join-Path -Path $scriptRoot -ChildPath "SmartMediaSorter_$runTimestamp.log"
}

$logDir = Split-Path -Path $LogPath -Parent
if ($logDir -and -not (Test-Path -LiteralPath $logDir -PathType Container)) {
    New-Item -Path $logDir -ItemType Directory -Force -WhatIf:$false | Out-Null
}
if (-not (Test-Path -LiteralPath $LogPath -PathType Leaf)) {
    New-Item -Path $LogPath -ItemType File -Force -WhatIf:$false | Out-Null
}

if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    Write-SorterLog -Message "Source directory does not exist: '$SourcePath'" -Level "ERROR" -LogFile $LogPath
    throw "Source directory does not exist: '$SourcePath'"
}

$resolvedSource = Get-NormalizedPath -Path (Resolve-Path -LiteralPath $SourcePath).Path
$normalizedDest = Get-NormalizedPath -Path $DestinationPath
$normalizedExtensions = @($IncludeExtensions | ForEach-Object { $_.ToLowerInvariant() })

if (Test-IsRecycleBinOrTrashPath -Path $resolvedSource) {
    Write-SorterLog -Message "Refusing to run: source path '$resolvedSource' is inside a Recycle Bin or trash directory." -Level "ERROR" -LogFile $LogPath
    throw "Refusing to run: source path '$resolvedSource' is inside a Recycle Bin or trash directory."
}
if (Test-IsRecycleBinOrTrashPath -Path $normalizedDest) {
    Write-SorterLog -Message "Refusing to run: destination path '$normalizedDest' is inside a Recycle Bin or trash directory." -Level "ERROR" -LogFile $LogPath
    throw "Refusing to run: destination path '$normalizedDest' is inside a Recycle Bin or trash directory."
}

if (-not $NoDatabase -and [string]::IsNullOrWhiteSpace($DatabasePath)) {
    $DatabasePath = Join-Path -Path $normalizedDest -ChildPath ".smartmediasorter_db.json"
}
elseif (-not [string]::IsNullOrWhiteSpace($DatabasePath)) {
    $DatabasePath = Get-NormalizedPath -Path $DatabasePath
}

$operationMode = if ($CopyOnly) { "COPY" } else { "MOVE" }
Write-SorterLog -Message "Starting SmartMediaSorter (Mode=$operationMode, WhatIf=$WhatIfPreference, Source='$resolvedSource', Destination='$normalizedDest', UseDatabase=$(-not $NoDatabase))" -Level "INFO" -LogFile $LogPath

# Load or initialize database
$sorterDb = if (-not $NoDatabase -and -not $RebuildDatabase -and -not [string]::IsNullOrWhiteSpace($DatabasePath)) {
    Import-SorterDatabase -DbFilePath $DatabasePath -LogFile $LogPath
}
else {
    New-EmptySorterDatabase
}

if ($RebuildDatabase -and -not $NoDatabase) {
    Write-SorterLog -Message "RebuildDatabase switch specified. Starting with a fresh hash database." -Level "INFO" -LogFile $LogPath
}

# Enumerate matching media files in SourcePath (skipping Recycle Bin, trash directories, and reparse points)
$allSourceFilesList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$dirStack = [System.Collections.Generic.Stack[System.IO.DirectoryInfo]]::new()
$dirStack.Push([System.IO.DirectoryInfo]::new($resolvedSource))

while ($dirStack.Count -gt 0) {
    $currDir = $dirStack.Pop()
    if (Test-IsRecycleBinOrTrashDirectoryName -DirectoryName $currDir.Name) {
        continue
    }

    try {
        $rawFiles = $currDir.GetFiles()
        foreach ($f in $rawFiles) {
            if ([bool]($f.Attributes -band ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)) -or
                $f.Name.StartsWith('.')) {
                continue
            }
            if ($normalizedExtensions.Count -eq 0 -or ($normalizedExtensions -contains $f.Extension.ToLowerInvariant())) {
                $allSourceFilesList.Add($f)
            }
        }
    }
    catch {
        Write-SorterLog -Message "Cannot list files in '$($currDir.FullName)': $($_.Exception.Message)" -Level "WARN" -LogFile $LogPath
    }

    if ($Recurse) {
        try {
            $subDirs = $currDir.GetDirectories()
            for ($idx = $subDirs.Length - 1; $idx -ge 0; $idx--) {
                $sub = $subDirs[$idx]
                if (Test-IsRecycleBinOrTrashDirectoryName -DirectoryName $sub.Name) {
                    continue
                }
                if ([bool]($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                    continue
                }
                $dirStack.Push($sub)
            }
        }
        catch {
            Write-SorterLog -Message "Cannot list subdirectories in '$($currDir.FullName)': $($_.Exception.Message)" -Level "WARN" -LogFile $LogPath
        }
    }
}

$allSourceFiles = @($allSourceFilesList | Sort-Object -Property FullName)

$stats = [ordered]@{
    TotalScanned      = $allSourceFiles.Count
    Moved             = 0
    Copied            = 0
    DuplicatesSkipped = 0
    FallbackDateUsed  = 0
    CacheHits         = 0
    DirectoryHashHits = 0
    Errors            = 0
    LogPath           = (Resolve-Path -LiteralPath $LogPath).Path
    DatabasePath      = if ($NoDatabase) { $null } else { $DatabasePath }
}

if ($allSourceFiles.Count -eq 0) {
    Write-SorterLog -Message "No matching image files found in '$resolvedSource'." -Level "WARN" -LogFile $LogPath
    if (-not $NoDatabase -and -not $WhatIfPreference -and -not [string]::IsNullOrWhiteSpace($DatabasePath)) {
        $sorterDb["SourceDirectories"][$resolvedSource] = @{
            DirectoryHash = (Get-DirectoryStateHash -Files @())
            FileCount     = 0
            UpdatedUtc    = [datetime]::UtcNow.ToString("o")
        }
        Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $LogPath
    }
    if ($PassThru) {
        return [PSCustomObject]$stats
    }
    return
}

# Group source files by directory and evaluate source directory state hashes
$sourceFilesByDir = @{}
foreach ($f in $allSourceFiles) {
    $dirNorm = Get-NormalizedPath -Path $f.DirectoryName
    if (-not $sourceFilesByDir.ContainsKey($dirNorm)) {
        $sourceFilesByDir[$dirNorm] = [System.Collections.Generic.List[object]]::new()
    }
    $sourceFilesByDir[$dirNorm].Add($f)
}

$unchangedSourceDirs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$sourceDirsToRefresh = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$destDirsToRefresh = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$currentSourceFileSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($dirKey in $sourceFilesByDir.Keys) {
    $dirFiles = $sourceFilesByDir[$dirKey]
    $computedDirHash = Get-DirectoryStateHash -Files $dirFiles
    if ($sorterDb["SourceDirectories"].ContainsKey($dirKey)) {
        $prevDirEntry = $sorterDb["SourceDirectories"][$dirKey]
        if ($prevDirEntry.ContainsKey("DirectoryHash") -and $prevDirEntry["DirectoryHash"] -eq $computedDirHash) {
            [void]$unchangedSourceDirs.Add($dirKey)
            $stats.DirectoryHashHits++
        }
    }
    [void]$sourceDirsToRefresh.Add($dirKey)
}

# Hash cache and in-memory target registry
$hashCache = @{}
$sessionMonthIndex = @{}
$simulatedTargetsByPrefix = @{}

# Determine which source files already have valid cached metadata in the database
$filesNeedingExif = [System.Collections.Generic.List[object]]::new()
$cachedFileDates = @{}

foreach ($file in $allSourceFiles) {
    $normSourceFile = Get-NormalizedPath -Path $file.FullName
    [void]$currentSourceFileSet.Add($normSourceFile)

    if ($sorterDb["SourceFiles"].ContainsKey($normSourceFile)) {
        $cachedEntry = $sorterDb["SourceFiles"][$normSourceFile]
        if ($cachedEntry["Length"] -eq $file.Length -and $cachedEntry["LastWriteTimeUtcTicks"] -eq $file.LastWriteTimeUtc.Ticks) {
            if ($cachedEntry.ContainsKey("PhotoDateIso") -and -not [string]::IsNullOrWhiteSpace($cachedEntry["PhotoDateIso"])) {
                $parsedCachedDate = ConvertFrom-ExifDateString -RawValue $cachedEntry["PhotoDateIso"]
                if ($null -ne $parsedCachedDate) {
                    $cachedFileDates[$normSourceFile] = [PSCustomObject]@{
                        Date   = $parsedCachedDate
                        Source = [string]$cachedEntry["DateSource"]
                    }
                    $stats.CacheHits++
                    continue
                }
            }
        }
    }

    $filesNeedingExif.Add($file)
}

# Prune stale source file entries that no longer exist in SourcePath
$staleSourceKeys = @(
    $sorterDb["SourceFiles"].Keys | Where-Object { -not $currentSourceFileSet.Contains($_) }
)
foreach ($staleKey in $staleSourceKeys) {
    $sorterDb["SourceFiles"].Remove($staleKey)
}

$exifMap = @{}
if ($filesNeedingExif.Count -gt 0) {
    $resolvedExifTool = Resolve-ExifToolBinary -ConfiguredPath $ExifToolPath
    if ($null -ne $resolvedExifTool) {
        $targetPathsForExif = @($filesNeedingExif | ForEach-Object { $_.FullName })
        Write-SorterLog -Message "Extracting EXIF metadata via ExifTool for $($filesNeedingExif.Count) new/modified file(s) (Cached=$($cachedFileDates.Count))..." -Level "INFO" -LogFile $LogPath
        $exifMap = Get-BatchExifMetadata -BinaryPath $resolvedExifTool -TargetPaths $targetPathsForExif -Recursive $false
    }
    else {
        Write-SorterLog -Message "ExifTool binary ('$ExifToolPath') is unavailable or failed to initialize. Using native .NET EXIF reader with LastWriteTime fallback." -Level "WARN" -LogFile $LogPath
    }
}
else {
    Write-SorterLog -Message "All $($allSourceFiles.Count) source file(s) matched database cache. Skipping ExifTool scan." -Level "INFO" -LogFile $LogPath
}

foreach ($file in $allSourceFiles) {
    try {
        $normSourceFile = Get-NormalizedPath -Path $file.FullName
        $normSourceDir = Get-NormalizedPath -Path $file.DirectoryName

        $photoDate = $null
        $dateTagSource = $null

        # 0. Use cached EXIF/fallback date from database if file is unchanged
        if ($cachedFileDates.ContainsKey($normSourceFile)) {
            $photoDate = $cachedFileDates[$normSourceFile].Date
            $dateTagSource = $cachedFileDates[$normSourceFile].Source
            if ($dateTagSource -eq "Fallback:LastWriteTime") {
                $stats.FallbackDateUsed++
            }
        }
        else {
            $exifEntry = if ($exifMap.ContainsKey($normSourceFile)) { $exifMap[$normSourceFile] } else { $null }

            # 1. Try ExifTool batch metadata first
            if ($null -ne $exifEntry) {
                foreach ($tagName in @("DateTimeOriginal", "CreateDate", "MediaCreateDate", "TrackCreateDate")) {
                    if ($null -ne $exifEntry.PSObject.Properties[$tagName]) {
                        $candidateDate = ConvertFrom-ExifDateString -RawValue $exifEntry.$tagName
                        if ($null -ne $candidateDate) {
                            $photoDate = $candidateDate
                            $dateTagSource = "EXIF:$tagName"
                            break
                        }
                    }
                }
            }

            # 2. Try native .NET EXIF reader if ExifTool did not provide a date
            if ($null -eq $photoDate) {
                $dotNetExif = Get-DotNetExifDate -FilePath $file.FullName
                if ($null -ne $dotNetExif) {
                    $photoDate = $dotNetExif.Date
                    $dateTagSource = $dotNetExif.Source
                }
            }

            # 3. Fallback to file LastWriteTime
            if ($null -eq $photoDate) {
                $photoDate = $file.LastWriteTime
                $dateTagSource = "Fallback:LastWriteTime"
                $stats.FallbackDateUsed++
                Write-SorterLog -Message "No EXIF date found for '$($file.FullName)'. Using LastWriteTime ($($photoDate.ToString('yyyy-MM-dd HH:mm:ss')))." -Level "FALLBACK" -LogFile $LogPath
            }
        }

        $yearFolder = $photoDate.ToString("yyyy")
        $monthFolder = $photoDate.ToString("yyyy-MM")
        $targetMonthDir = Join-Path -Path $normalizedDest -ChildPath (Join-Path -Path $yearFolder -ChildPath $monthFolder)
        $normTargetMonthDir = Get-NormalizedPath -Path $targetMonthDir

        $timestampPrefix = $photoDate.ToString("yyyy-MM-dd_HHmmss")
        $extLower = $file.Extension.ToLowerInvariant()
        $bucketKey = ("{0}\{1}{2}" -f $normTargetMonthDir, $timestampPrefix, $extLower).ToLowerInvariant()

        # Gather existing files in target month directory via synchronized database cache
        $monthAllFiles = Sync-DestinationMonthDirectory `
            -MonthDirPath $normTargetMonthDir `
            -Database $sorterDb `
            -SessionMonthIndex $sessionMonthIndex `
            -HashCache $hashCache `
            -Stats $stats

        $existingCandidates = [System.Collections.Generic.List[object]]::new()
        $expectedPrefixWithUnderscore = "{0}_" -f $timestampPrefix

        foreach ($mf in $monthAllFiles) {
            if ($mf.Name.StartsWith($expectedPrefixWithUnderscore, [System.StringComparison]::OrdinalIgnoreCase) -and
                $mf.Name.EndsWith($extLower, [System.StringComparison]::OrdinalIgnoreCase)) {
                $existingCandidates.Add($mf)
            }
        }

        # Include files already planned/placed in this bucket during the current session (e.g., under -WhatIf)
        if ($simulatedTargetsByPrefix.ContainsKey($bucketKey)) {
            foreach ($simItem in $simulatedTargetsByPrefix[$bucketKey]) {
                $alreadyListed = $existingCandidates | Where-Object {
                    (Get-NormalizedPath -Path $_.FullName) -eq (Get-NormalizedPath -Path $simItem.DestinationFullName)
                }
                if (-not $alreadyListed) {
                    $existingCandidates.Add([PSCustomObject]@{
                        FullName   = $simItem.DestinationFullName
                        Name       = $simItem.DestinationName
                        Length     = $simItem.Length
                        SourcePath = $simItem.SourceFullName
                    })
                }
            }
        }

        # Check for duplicates among files sharing the same timestamp and extension
        $isDuplicate = $false
        $duplicateTargetPath = $null

        if ($existingCandidates.Count -gt 0) {
            foreach ($candidate in $existingCandidates) {
                if ((Get-NormalizedPath -Path $candidate.FullName) -eq $normSourceFile) {
                    $isDuplicate = $true
                    $duplicateTargetPath = $candidate.FullName
                    break
                }

                # Fast pre-check by byte size before computing/looking up SHA-256
                if ($candidate.Length -eq $file.Length) {
                    $sourceHash = Get-CachedFileHash -FilePath $file.FullName -Cache $hashCache -Database $sorterDb -Stats $stats
                    $hashLookupPath = if (Test-Path -LiteralPath $candidate.FullName -PathType Leaf) {
                        $candidate.FullName
                    } elseif ($null -ne $candidate.PSObject.Properties["SourcePath"] -and $candidate.SourcePath) {
                        $candidate.SourcePath
                    } else {
                        $candidate.FullName
                    }
                    $candidateHash = Get-CachedFileHash -FilePath $hashLookupPath -Cache $hashCache -Database $sorterDb -Stats $stats

                    if ($sourceHash -eq $candidateHash) {
                        $isDuplicate = $true
                        $duplicateTargetPath = $candidate.FullName
                        break
                    }
                }
            }
        }

        if ($isDuplicate) {
            $stats.DuplicatesSkipped++
            $knownSha = if ($hashCache.ContainsKey($normSourceFile)) { $hashCache[$normSourceFile] } else { $null }
            $sorterDb["SourceFiles"][$normSourceFile] = @{
                DirectoryPath         = $normSourceDir
                Length                = [long]$file.Length
                LastWriteTimeUtcTicks = [long]$file.LastWriteTimeUtc.Ticks
                PhotoDateIso          = $photoDate.ToString("yyyy-MM-dd HH:mm:ss")
                DateSource            = $dateTagSource
                Sha256                = $knownSha
                LastAction            = "Duplicate"
                TargetPath            = (Get-NormalizedPath -Path $duplicateTargetPath)
            }
            Write-SorterLog -Message "Duplicate detected: '$($file.FullName)' matches existing '$duplicateTargetPath'. Skipping (left in source)." -Level "DUPLICATE" -LogFile $LogPath
            continue
        }

        # Determine the lowest available 3-digit sequence number (_000, _001, _002, ...)
        $occupiedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($c in $existingCandidates) {
            [void]$occupiedNames.Add($c.Name)
        }

        $seqIndex = 0
        $targetFileName = $null
        $targetFilePath = $null

        while ($true) {
            $candidateName = "{0}_{1:D3}{2}" -f $timestampPrefix, $seqIndex, $extLower
            $candidatePath = Join-Path -Path $normTargetMonthDir -ChildPath $candidateName

            if (-not $occupiedNames.Contains($candidateName) -and -not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                $targetFileName = $candidateName
                $targetFilePath = $candidatePath
                break
            }
            $seqIndex++
        }

        $normTarget = Get-NormalizedPath -Path $targetFilePath

        # Record in session bucket so subsequent files in the same second see this slot
        if (-not $simulatedTargetsByPrefix.ContainsKey($bucketKey)) {
            $simulatedTargetsByPrefix[$bucketKey] = [System.Collections.Generic.List[object]]::new()
        }
        $simulatedTargetsByPrefix[$bucketKey].Add([PSCustomObject]@{
            DestinationFullName = $normTarget
            DestinationName     = $targetFileName
            SourceFullName      = $file.FullName
            Length              = $file.Length
        })

        # Ensure destination directory exists and perform Move/Copy
        $actionVerb = if ($CopyOnly) { "Copy photo ($dateTagSource)" } else { "Move photo ($dateTagSource)" }
        if ($PSCmdlet.ShouldProcess("$($file.FullName) -> $normTarget", $actionVerb)) {
            if (-not (Test-Path -LiteralPath $normTargetMonthDir -PathType Container)) {
                New-Item -Path $normTargetMonthDir -ItemType Directory -Force | Out-Null
            }

            # Compute SHA-256 before/during transfer so the database always has the file hash ready
            $fileSha256 = Get-CachedFileHash -FilePath $file.FullName -Cache $hashCache -Database $sorterDb -Stats $stats
            $hashCache[$normTarget] = $fileSha256

            if ($CopyOnly) {
                Copy-Item -LiteralPath $file.FullName -Destination $normTarget -Force
                $copiedItem = Get-Item -LiteralPath $normTarget

                $sorterDb["SourceFiles"][$normSourceFile] = @{
                    DirectoryPath         = $normSourceDir
                    Length                = [long]$file.Length
                    LastWriteTimeUtcTicks = [long]$file.LastWriteTimeUtc.Ticks
                    PhotoDateIso          = $photoDate.ToString("yyyy-MM-dd HH:mm:ss")
                    DateSource            = $dateTagSource
                    Sha256                = $fileSha256
                    LastAction            = "Copied"
                    TargetPath            = $normTarget
                }

                $sorterDb["DestinationFiles"][$normTarget] = @{
                    Name                  = $targetFileName
                    DirectoryPath         = $normTargetMonthDir
                    Length                = [long]$copiedItem.Length
                    LastWriteTimeUtcTicks = [long]$copiedItem.LastWriteTimeUtc.Ticks
                    Sha256                = $fileSha256
                }

                $monthAllFiles.Add([PSCustomObject]@{
                    FullName = $normTarget
                    Name     = $targetFileName
                    Length   = [long]$copiedItem.Length
                })

                [void]$destDirsToRefresh.Add($normTargetMonthDir)
                $stats.Copied++
                Write-SorterLog -Message "Copied '$($file.FullName)' -> '$normTarget' [$dateTagSource]" -Level "COPIED" -LogFile $LogPath
            }
            else {
                Move-Item -LiteralPath $file.FullName -Destination $normTarget -Force
                $movedItem = Get-Item -LiteralPath $normTarget

                $sorterDb["SourceFiles"].Remove($normSourceFile)
                $sorterDb["DestinationFiles"][$normTarget] = @{
                    Name                  = $targetFileName
                    DirectoryPath         = $normTargetMonthDir
                    Length                = [long]$movedItem.Length
                    LastWriteTimeUtcTicks = [long]$movedItem.LastWriteTimeUtc.Ticks
                    Sha256                = $fileSha256
                }

                $monthAllFiles.Add([PSCustomObject]@{
                    FullName = $normTarget
                    Name     = $targetFileName
                    Length   = [long]$movedItem.Length
                })

                [void]$sourceDirsToRefresh.Add($normSourceDir)
                [void]$destDirsToRefresh.Add($normTargetMonthDir)
                $stats.Moved++
                Write-SorterLog -Message "Moved '$($file.FullName)' -> '$normTarget' [$dateTagSource]" -Level "MOVED" -LogFile $LogPath
            }
        }
        else {
            Write-SorterLog -Message "WhatIf: Would $($operationMode.ToLowerInvariant()) '$($file.FullName)' -> '$normTarget' [$dateTagSource]" -Level "WHATIF" -LogFile $LogPath
        }
    }
    catch {
        $stats.Errors++
        Write-SorterLog -Message "Error processing '$($file.FullName)': $($_.Exception.Message)" -Level "ERROR" -LogFile $LogPath
    }
}

# Persist updated directory and file hashes to database (skip when running under -WhatIf)
if (-not $NoDatabase -and -not $WhatIfPreference -and -not [string]::IsNullOrWhiteSpace($DatabasePath)) {
    [void]$sourceDirsToRefresh.Add($resolvedSource)
    Update-DirectoryHashesOnDisk `
        -Database $sorterDb `
        -SourceDirsToRefresh $sourceDirsToRefresh `
        -DestDirsToRefresh $destDirsToRefresh `
        -NormalizedExtensions $normalizedExtensions

    Export-SorterDatabase -Database $sorterDb -DbFilePath $DatabasePath -LogFile $LogPath
    Write-SorterLog -Message "Updated hash database at '$DatabasePath' (SourceDirs=$($sorterDb['SourceDirectories'].Count), DestDirs=$($sorterDb['DestinationDirectories'].Count), DestFiles=$($sorterDb['DestinationFiles'].Count))." -Level "INFO" -LogFile $LogPath
}

$summaryLine = "Completed SmartMediaSorter: Scanned=$($stats.TotalScanned), Moved=$($stats.Moved), Copied=$($stats.Copied), DuplicatesSkipped=$($stats.DuplicatesSkipped), FallbackDateUsed=$($stats.FallbackDateUsed), CacheHits=$($stats.CacheHits), DirHashHits=$($stats.DirectoryHashHits), Errors=$($stats.Errors). Log='$($stats.LogPath)'"
Write-SorterLog -Message $summaryLine -Level "INFO" -LogFile $LogPath

Write-Host ""
Write-Host "=== SmartMediaSorter Summary ===" -ForegroundColor Cyan
Write-Host "Total scanned       : $($stats.TotalScanned)"
Write-Host "Moved               : $($stats.Moved)" -ForegroundColor Green
Write-Host "Copied              : $($stats.Copied)" -ForegroundColor Cyan
Write-Host "Duplicates skipped  : $($stats.DuplicatesSkipped)" -ForegroundColor Yellow
Write-Host "Fallback date used  : $($stats.FallbackDateUsed)"
Write-Host "Database cache hits : $($stats.CacheHits) (DirHash hits: $($stats.DirectoryHashHits))" -ForegroundColor Gray
Write-Host "Errors              : $($stats.Errors)" -ForegroundColor $(if ($stats.Errors -gt 0) { "Red" } else { "Gray" })
Write-Host "Database file       : $(if ($NoDatabase) { '(disabled)' } else { $stats.DatabasePath })"
Write-Host "Log file            : $($stats.LogPath)"
Write-Host "================================" -ForegroundColor Cyan

if ($PassThru) {
    return [PSCustomObject]$stats
}

#endregion Main Execution
