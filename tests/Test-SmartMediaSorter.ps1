#Requires -Version 7.0

<#
.SYNOPSIS
    Automated integration tests for SmartMediaSorter.ps1.

.DESCRIPTION
    Creates temporary source and destination directories, generates sample JPEG files
    with EXIF DateTimeOriginal (0x9003) metadata, and verifies:
    1. -WhatIf simulation mode (no files moved, no database persisted).
    2. Standard Move mode with EXIF DateTimeOriginal -> \YYYY\YYYY-MM\YYYY-MM-DD_HHMMSS_000.jpg
       and creation of persistent hash database (.smartmediasorter_db.json).
    3. Burst-mode sequence numbering (_001.jpg) for a different photo in the same second.
    4. Duplicate detection via SHA-256 (duplicate stays in SourcePath).
    5. Fallback to LastWriteTime when EXIF date metadata is missing.
    6. Database cache and directory hash hits on subsequent runs, plus hash updates
       in source and destination directories when new files are moved.
    7. -CopyOnly, -RebuildDatabase, and -NoDatabase modes.
#>

[CmdletBinding()]
param(
    [string]$ExifToolPath = "exiftool"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
$sorterScript = Join-Path -Path $projectRoot -ChildPath "SmartMediaSorter.ps1"
$dupFinderScript = Join-Path -Path $projectRoot -ChildPath "Find-MediaDuplicates.ps1"
$removerScript = Join-Path -Path $projectRoot -ChildPath "Remove-DuplicateFolders.ps1"

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) {
        throw "ASSERTION FAILED: $Message"
    }
    Write-Host "[PASS] $Message" -ForegroundColor Green
}

function New-SampleJpeg {
    <#
    .SYNOPSIS
        Generates a valid JPEG image with a specified solid color and optional EXIF DateTimeOriginal (0x9003).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][int]$Red,
        [Parameter(Mandatory = $true)][int]$Green,
        [Parameter(Mandatory = $true)][int]$Blue,
        [Parameter()][string]$ExifDateTimeOriginal = ""
    )

    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $color = [System.Drawing.Color]::FromArgb($Red, $Green, $Blue)
        $brush = New-Object System.Drawing.SolidBrush($color)
        $graphics.FillRectangle($brush, 0, 0, 32, 32)
        $brush.Dispose()

        $bmp.Save($FilePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    finally {
        $graphics.Dispose()
        $bmp.Dispose()
    }

    # Inject standard EXIF 2.3 APP1 segment (IFD0 -> ExifOffset 0x8769 -> DateTimeOriginal 0x9003)
    if (-not [string]::IsNullOrWhiteSpace($ExifDateTimeOriginal)) {
        $rawJpeg = [System.IO.File]::ReadAllBytes($FilePath)
        $dateBytes = [System.Text.Encoding]::ASCII.GetBytes($ExifDateTimeOriginal.Substring(0, 19) + [char]0)

        [byte[]]$app1Header = @(
            0xFF, 0xE1,                         # APP1 Marker
            0x00, 0x48,                         # Length = 72 bytes (2 + 6 + 64)
            0x45, 0x78, 0x69, 0x66, 0x00, 0x00, # "Exif\0\0"
            # TIFF Header (8 bytes, Little-Endian)
            0x49, 0x49, 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
            # IFD0 at offset 8 (18 bytes): 1 entry -> ExifOffset (0x8769) pointing to offset 26 (0x1A)
            0x01, 0x00,
            0x69, 0x87, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x1A, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            # Exif SubIFD at offset 26 (18 bytes): 1 entry -> DateTimeOriginal (0x9003) pointing to offset 44 (0x2C)
            0x01, 0x00,
            0x03, 0x90, 0x02, 0x00, 0x14, 0x00, 0x00, 0x00, 0x2C, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        )

        $outStream = New-Object System.IO.MemoryStream
        try {
            # Write SOI (0xFF, 0xD8)
            $outStream.Write($rawJpeg, 0, 2)
            # Write APP1 EXIF segment
            $outStream.Write($app1Header, 0, $app1Header.Length)
            $outStream.Write($dateBytes, 0, 20)
            # Write rest of JPEG
            $outStream.Write($rawJpeg, 2, $rawJpeg.Length - 2)
            [System.IO.File]::WriteAllBytes($FilePath, $outStream.ToArray())
        }
        finally {
            $outStream.Dispose()
        }
    }
}

$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "SmartMediaSorter_Tests_$PID"
if (Test-Path -LiteralPath $testRoot) {
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}

$sourceDir = Join-Path -Path $testRoot -ChildPath "Photos"
$subSourceDir = Join-Path -Path $sourceDir -ChildPath "Vacation"
$destDir = Join-Path -Path $testRoot -ChildPath "Pictures"
$logFile = Join-Path -Path $testRoot -ChildPath "test_run.log"
$expectedDbPath = Join-Path -Path $destDir -ChildPath ".smartmediasorter_db.json"

New-Item -Path $subSourceDir -ItemType Directory -Force | Out-Null
New-Item -Path $destDir -ItemType Directory -Force | Out-Null

try {
    Write-Host "Preparing test images in '$sourceDir'..." -ForegroundColor Cyan

    # 1. Primary photo with EXIF DateTimeOriginal = 2015:08:30 11:40:23
    $img1 = Join-Path -Path $sourceDir -ChildPath "IMG_0001.JPG"
    New-SampleJpeg -FilePath $img1 -Red 220 -Green 20 -Blue 60 -ExifDateTimeOriginal "2015:08:30 11:40:23"

    # 2. Different photo taken in the exact same second (2015:08:30 11:40:23) -> should become _001.jpg
    $img2 = Join-Path -Path $subSourceDir -ChildPath "IMG_0002.JPG"
    New-SampleJpeg -FilePath $img2 -Red 30 -Green 144 -Blue 255 -ExifDateTimeOriginal "2015:08:30 11:40:23"

    # 3. Exact duplicate of IMG_0001.JPG (identical bytes & SHA-256) -> should be skipped and remain in source
    $imgDup = Join-Path -Path $subSourceDir -ChildPath "IMG_0001_copy.JPG"
    Copy-Item -LiteralPath $img1 -Destination $imgDup -Force

    # 4. Photo without EXIF tags -> should fall back to LastWriteTime (2021-12-05 09:15:07)
    $imgNoExif = Join-Path -Path $sourceDir -ChildPath "NO_EXIF.jpg"
    New-SampleJpeg -FilePath $imgNoExif -Red 50 -Green 205 -Blue 50
    $fallbackDate = [datetime]::ParseExact("2021-12-05 09:15:07", "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    (Get-Item -LiteralPath $imgNoExif).LastWriteTime = $fallbackDate

    # --- TEST 1: -WhatIf simulation ---
    Write-Host "`nRunning Test 1: -WhatIf Mode..." -ForegroundColor Cyan
    $whatIfStats = & $sorterScript -SourcePath $sourceDir -DestinationPath $destDir -ExifToolPath $ExifToolPath -LogPath $logFile -WhatIf -PassThru
    Assert-True ($whatIfStats.TotalScanned -eq 4) "WhatIf scanned all 4 source files"
    Assert-True ($whatIfStats.Moved -eq 0) "WhatIf did not move any files"
    Assert-True ($whatIfStats.DuplicatesSkipped -eq 1) "WhatIf accurately detected 1 duplicate within batch"
    Assert-True ($whatIfStats.FallbackDateUsed -eq 1) "WhatIf used fallback date only for the 1 file without EXIF"
    Assert-True (@(Get-ChildItem -LiteralPath $destDir -Recurse -File).Count -eq 0) "Destination directory is still empty after -WhatIf"
    Assert-True (-not (Test-Path -LiteralPath $expectedDbPath -PathType Leaf)) "Database file was not created during -WhatIf"

    # --- TEST 2: Default Move Mode & Database Creation ---
    Write-Host "`nRunning Test 2: Default Move Mode & Database Creation..." -ForegroundColor Cyan
    $moveStats = & $sorterScript -SourcePath $sourceDir -DestinationPath $destDir -ExifToolPath $ExifToolPath -LogPath $logFile -PassThru

    Assert-True ($moveStats.TotalScanned -eq 4) "Scanned 4 files"
    Assert-True ($moveStats.Moved -eq 3) "Moved 3 unique files"
    Assert-True ($moveStats.DuplicatesSkipped -eq 1) "Skipped 1 duplicate file"
    Assert-True ($moveStats.FallbackDateUsed -eq 1) "Used LastWriteTime fallback for 1 file without EXIF"
    Assert-True ($moveStats.Errors -eq 0) "Zero errors occurred"
    Assert-True (Test-Path -LiteralPath $expectedDbPath -PathType Leaf) "Persistent database .smartmediasorter_db.json was created"

    $expectedFile000 = Join-Path -Path $destDir -ChildPath "2015\2015-08\2015-08-30_114023_000.jpg"
    $expectedFile001 = Join-Path -Path $destDir -ChildPath "2015\2015-08\2015-08-30_114023_001.jpg"
    $expectedFallback = Join-Path -Path $destDir -ChildPath "2021\2021-12\2021-12-05_091507_000.jpg"

    Assert-True (Test-Path -LiteralPath $expectedFile000 -PathType Leaf) "Created 2015\2015-08\2015-08-30_114023_000.jpg"
    Assert-True (Test-Path -LiteralPath $expectedFile001 -PathType Leaf) "Created 2015\2015-08\2015-08-30_114023_001.jpg for second photo in same second"
    Assert-True (Test-Path -LiteralPath $expectedFallback -PathType Leaf) "Created 2021\2021-12\2021-12-05_091507_000.jpg using LastWriteTime fallback"

    $dbAfterMove = Get-Content -LiteralPath $expectedDbPath -Raw | ConvertFrom-Json
    $subSourceHashBefore = $dbAfterMove.SourceDirectories.$subSourceDir.DirectoryHash
    $dest2015MonthDir = Join-Path -Path $destDir -ChildPath "2015\2015-08"
    $dest2015HashBefore = $dbAfterMove.DestinationDirectories.$dest2015MonthDir.DirectoryHash

    Assert-True (-not [string]::IsNullOrWhiteSpace($subSourceHashBefore)) "Source subdirectory hash recorded in database"
    Assert-True (-not [string]::IsNullOrWhiteSpace($dest2015HashBefore)) "Destination month directory hash recorded in database"

    # --- TEST 3: Re-running against remaining duplicate in SourcePath (Cache & DirHash Hits) ---
    Write-Host "`nRunning Test 3: Re-run against existing destination (Database Cache Hits)..." -ForegroundColor Cyan
    $rerunStats = & $sorterScript -SourcePath $sourceDir -DestinationPath $destDir -ExifToolPath $ExifToolPath -LogPath $logFile -PassThru
    Assert-True ($rerunStats.Moved -eq 0) "Zero files moved on re-run"
    Assert-True ($rerunStats.DuplicatesSkipped -eq 1) "Remaining duplicate detected against existing destination file"
    Assert-True ($rerunStats.CacheHits -ge 2) "Database cache hits recorded for source metadata and SHA-256 hashes (CacheHits=$($rerunStats.CacheHits))"
    Assert-True ($rerunStats.DirectoryHashHits -ge 2) "Directory hash hits recorded for unchanged source and destination folders (DirHashHits=$($rerunStats.DirectoryHashHits))"

    # --- TEST 4: Adding a new photo updates source and destination directory hashes ---
    Write-Host "`nRunning Test 4: Moving a new photo updates source and destination directory hashes..." -ForegroundColor Cyan
    $imgNewInSub = Join-Path -Path $subSourceDir -ChildPath "IMG_0003_NEW.JPG"
    New-SampleJpeg -FilePath $imgNewInSub -Red 128 -Green 0 -Blue 128 -ExifDateTimeOriginal "2015:08:30 11:40:23"

    $incrementalStats = & $sorterScript -SourcePath $sourceDir -DestinationPath $destDir -ExifToolPath $ExifToolPath -LogPath $logFile -PassThru
    $expectedFile002 = Join-Path -Path $destDir -ChildPath "2015\2015-08\2015-08-30_114023_002.jpg"

    Assert-True ($incrementalStats.Moved -eq 1) "Moved 1 newly added photo"
    Assert-True ($incrementalStats.DuplicatesSkipped -eq 1) "Still skipped the 1 known duplicate"
    Assert-True (Test-Path -LiteralPath $expectedFile002 -PathType Leaf) "Created _002.jpg in destination 2015-08 folder"

    $dbAfterIncremental = Get-Content -LiteralPath $expectedDbPath -Raw | ConvertFrom-Json
    $dest2015HashAfter = $dbAfterIncremental.DestinationDirectories.$dest2015MonthDir.DirectoryHash
    Assert-True ($dest2015HashAfter -ne $dest2015HashBefore) "Destination month directory hash changed after new file arrived"
    Assert-True ($null -ne $dbAfterIncremental.DestinationFiles.$expectedFile002.Sha256) "New destination file SHA-256 hash stored in database"

    # --- TEST 5: -CopyOnly and -RebuildDatabase modes ---
    Write-Host "`nRunning Test 5: -CopyOnly & -RebuildDatabase modes..." -ForegroundColor Cyan
    $imgCopyTest = Join-Path -Path $sourceDir -ChildPath "IMG_COPY_MODE.jpg"
    New-SampleJpeg -FilePath $imgCopyTest -Red 255 -Green 215 -Blue 0 -ExifDateTimeOriginal "2018:05:14 18:30:00"

    $copyStats = & $sorterScript -SourcePath $sourceDir -DestinationPath $destDir -ExifToolPath $ExifToolPath -LogPath $logFile -CopyOnly -RebuildDatabase -PassThru
    $expectedCopiedFile = Join-Path -Path $destDir -ChildPath "2018\2018-05\2018-05-14_183000_000.jpg"

    Assert-True ($copyStats.Copied -eq 1) "Copied 1 new file in -CopyOnly mode"
    Assert-True ($copyStats.DirectoryHashHits -eq 0) "RebuildDatabase ignored previous directory hashes as expected"
    Assert-True (Test-Path -LiteralPath $expectedCopiedFile -PathType Leaf) "Destination file exists after -CopyOnly"
    Assert-True (Test-Path -LiteralPath $imgCopyTest -PathType Leaf) "Source file was preserved in -CopyOnly mode"

    # --- TEST 6: Find-MediaDuplicates.ps1 (Filename Hint + Content SHA-256 + Local _Duplicates Folder) ---
    Write-Host "`nRunning Test 6: Find-MediaDuplicates.ps1 (Name hint -> SHA-256 -> Local _Duplicates folder)..." -ForegroundColor Cyan
    $dupTestDir = Join-Path -Path $testRoot -ChildPath "DupScanRoot"
    $dupSubDir = Join-Path -Path $dupTestDir -ChildPath "AlbumA"
    New-Item -Path $dupSubDir -ItemType Directory -Force | Out-Null

    # Original file in AlbumA
    $origPhoto = Join-Path -Path $dupSubDir -ChildPath "DSC_5000.jpg"
    New-SampleJpeg -FilePath $origPhoto -Red 10 -Green 120 -Blue 240 -ExifDateTimeOriginal "2022:07:10 14:20:00"

    # Duplicate 1 in same folder with " (1)" suffix
    $dupPhoto1 = Join-Path -Path $dupSubDir -ChildPath "DSC_5000 (1).jpg"
    Copy-Item -LiteralPath $origPhoto -Destination $dupPhoto1 -Force

    # Duplicate 2 in same folder with " - kopia" suffix
    $dupPhoto2 = Join-Path -Path $dupSubDir -ChildPath "DSC_5000 - kopia.jpg"
    Copy-Item -LiteralPath $origPhoto -Destination $dupPhoto2 -Force

    # Different extension with identical bytes -> MUST NOT be treated as duplicate
    $diffExtFile = Join-Path -Path $dupSubDir -ChildPath "DSC_5000.png"
    Copy-Item -LiteralPath $origPhoto -Destination $diffExtFile -Force

    # Same normalized name ("DSC_5000 (2).jpg") and extension, but different image content -> MUST NOT be treated as duplicate
    $diffContentPhoto = Join-Path -Path $dupSubDir -ChildPath "DSC_5000 (2).jpg"
    New-SampleJpeg -FilePath $diffContentPhoto -Red 240 -Green 50 -Blue 20 -ExifDateTimeOriginal "2022:07:10 14:20:00"

    # Completely different filename ("RENAMED_BACKUP_99.jpg"), same extension (.jpg) and identical content -> Stage 2 duplicate
    $renamedDupPhoto = Join-Path -Path $dupSubDir -ChildPath "RENAMED_BACKUP_99.jpg"
    Copy-Item -LiteralPath $origPhoto -Destination $renamedDupPhoto -Force

    # Test 6a: -WhatIf simulation first
    $dupWhatIfStats = & $dupFinderScript -Path $dupTestDir -LogPath $logFile -WhatIf -PassThru
    Assert-True ($dupWhatIfStats.TotalScanned -eq 6) "Find-MediaDuplicates WhatIf scanned all 6 files"
    Assert-True ($dupWhatIfStats.DuplicateSetsFound -ge 1) "Find-MediaDuplicates WhatIf found duplicate sets"
    Assert-True ($dupWhatIfStats.DuplicatesMoved -eq 0) "Find-MediaDuplicates WhatIf did not move any files"
    Assert-True (Test-Path -LiteralPath $dupPhoto1 -PathType Leaf) "Duplicate 1 still in place after -WhatIf"

    # Test 6b: Actual execution moves duplicates to AlbumA\_Duplicates
    $dupRunStats = & $dupFinderScript -Path $dupTestDir -LogPath $logFile -PassThru
    $localDuplicatesFolder = Join-Path -Path $dupSubDir -ChildPath "_Duplicates"
    $movedDup1 = Join-Path -Path $localDuplicatesFolder -ChildPath "DSC_5000 (1).jpg"
    $movedDup2 = Join-Path -Path $localDuplicatesFolder -ChildPath "DSC_5000 - kopia.jpg"
    $movedRenamedDup = Join-Path -Path $localDuplicatesFolder -ChildPath "RENAMED_BACKUP_99.jpg"

    Assert-True ($dupRunStats.DuplicatesMoved -eq 3) "Moved all 3 confirmed duplicates (Stage 1 name-hint + Stage 2 content) to _Duplicates subfolder"
    Assert-True (Test-Path -LiteralPath $origPhoto -PathType Leaf) "Primary file DSC_5000.jpg was kept in AlbumA"
    Assert-True (Test-Path -LiteralPath $movedDup1 -PathType Leaf) "DSC_5000 (1).jpg moved to AlbumA\_Duplicates"
    Assert-True (Test-Path -LiteralPath $movedDup2 -PathType Leaf) "DSC_5000 - kopia.jpg moved to AlbumA\_Duplicates"
    Assert-True (Test-Path -LiteralPath $movedRenamedDup -PathType Leaf) "RENAMED_BACKUP_99.jpg moved to AlbumA\_Duplicates via Stage 2 content check"
    Assert-True (Test-Path -LiteralPath $diffExtFile -PathType Leaf) "File with different extension (DSC_5000.png) was NOT treated as duplicate"
    Assert-True (Test-Path -LiteralPath $diffContentPhoto -PathType Leaf) "File with matching name pattern but different content (DSC_5000 (2).jpg) was NOT treated as duplicate"

    # Test 6c: Re-running Find-MediaDuplicates ignores _Duplicates folder and uses cached hashes
    $dupRerunStats = & $dupFinderScript -Path $dupTestDir -LogPath $logFile -PassThru
    Assert-True ($dupRerunStats.TotalScanned -eq 3) "Re-run ignored files inside _Duplicates folder (scanned 3 remaining files)"
    Assert-True ($dupRerunStats.DuplicatesMoved -eq 0) "Zero duplicates moved on second run"
    Assert-True ($dupRerunStats.DirectoryHashHits -ge 1) "DirectoryHash hit recorded on unchanged AlbumA directory"

    # --- TEST 7: Stage 3 ReferenceDb check (leftover duplicates in Photos vs Pictures) & Thumbs.db exclusion ---
    Write-Host "`nRunning Test 7: Stage 3 ReferenceDb check (Photos vs Pictures database) & Thumbs.db exclusion..." -ForegroundColor Cyan
    $dummyThumbs = Join-Path -Path $subSourceDir -ChildPath "Thumbs.db"
    [System.IO.File]::WriteAllBytes($dummyThumbs, [byte[]]@(1, 2, 3, 4, 5, 6, 7, 8))

    $refCheckStats = & $dupFinderScript -Path $sourceDir -ReferencePath $destDir -LogPath $logFile -PassThru
    $quarantinedInSub = Join-Path -Path $subSourceDir -ChildPath "_Duplicates\IMG_0001_copy.JPG"
    $quarantinedInRoot = Join-Path -Path $sourceDir -ChildPath "_Duplicates\IMG_COPY_MODE.jpg"

    Assert-True ($refCheckStats.TotalScanned -eq 2) "Ignored Thumbs.db and scanned only the 2 remaining media files in Photos"
    Assert-True ($refCheckStats.DuplicatesMoved -eq 2) "Detected and moved both leftover files in Photos that matched originals in Pictures"
    Assert-True (Test-Path -LiteralPath $quarantinedInSub -PathType Leaf) "Moved Photos\Vacation\IMG_0001_copy.JPG -> Photos\Vacation\_Duplicates\IMG_0001_copy.JPG"
    Assert-True (Test-Path -LiteralPath $quarantinedInRoot -PathType Leaf) "Moved Photos\IMG_COPY_MODE.jpg -> Photos\_Duplicates\IMG_COPY_MODE.jpg"
    Assert-True (Test-Path -LiteralPath $dummyThumbs -PathType Leaf) "Thumbs.db was untouched"

    # --- TEST 8: Stage 1B Metadata-divergent JPEGs (-01 vs _000, sidecar .xmp priority, and burst preservation) ---
    Write-Host "`nRunning Test 8: Stage 1B ImagePayload hash (-01 vs _000 with different EXIF/XMP sizes & .xmp sidecars)..." -ForegroundColor Cyan
    $payloadTestDir = Join-Path -Path $testRoot -ChildPath "Pictures2025\2025-01"
    New-Item -Path $payloadTestDir -ItemType Directory -Force | Out-Null

    # Pair A: 2025-01-02_153726_000.jpg (smaller header) vs 2025-01-02_153726-01.jpg (extra 1024-byte APP1 XMP segment + .xmp sidecar)
    $pairA_000 = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153726_000.jpg"
    $pairA_dash01 = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153726-01.jpg"
    $pairA_xmp = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153726-01.jpg.xmp"
    New-SampleJpeg -FilePath $pairA_000 -Red 70 -Green 130 -Blue 180 -ExifDateTimeOriginal "2025:01:02 15:37:26"

    # Create -01.jpg with an extra APP1 XMP metadata block (1000 bytes larger) but identical JPEG image scan stream
    $rawA = [System.IO.File]::ReadAllBytes($pairA_000)
    $extraApp1 = [byte[]]::new(1004)
    $extraApp1[0] = 0xFF; $extraApp1[1] = 0xE1; $extraApp1[2] = 0x03; $extraApp1[3] = 0xEA # APP1 length = 1002
    $msA = New-Object System.IO.MemoryStream
    $msA.Write($rawA, 0, 2)
    $msA.Write($extraApp1, 0, $extraApp1.Length)
    $msA.Write($rawA, 2, $rawA.Length - 2)
    [System.IO.File]::WriteAllBytes($pairA_dash01, $msA.ToArray())
    $msA.Dispose()
    [System.IO.File]::WriteAllText($pairA_xmp, "<x:xmpmeta>Sidecar A</x:xmpmeta>")

    # Pair B: 2025-01-02_153731.jpg (larger header + .xmp sidecar) vs 2025-01-02_153731_000.jpg (smaller header) + genuine burst _001.jpg
    $pairB_plain = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153731.jpg"
    $pairB_000 = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153731_000.jpg"
    $pairB_xmp = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153731.jpg.xmp"
    $pairB_burst001 = Join-Path -Path $payloadTestDir -ChildPath "2025-01-02_153731_001.jpg"

    New-SampleJpeg -FilePath $pairB_000 -Red 200 -Green 90 -Blue 40 -ExifDateTimeOriginal "2025:01:02 15:37:31"
    $rawB = [System.IO.File]::ReadAllBytes($pairB_000)
    $msB = New-Object System.IO.MemoryStream
    $msB.Write($rawB, 0, 2)
    $msB.Write($extraApp1, 0, $extraApp1.Length)
    $msB.Write($rawB, 2, $rawB.Length - 2)
    [System.IO.File]::WriteAllBytes($pairB_plain, $msB.ToArray())
    $msB.Dispose()
    [System.IO.File]::WriteAllText($pairB_xmp, "<x:xmpmeta>Sidecar B</x:xmpmeta>")

    # Genuine burst shot in the same second (different visual content)
    New-SampleJpeg -FilePath $pairB_burst001 -Red 15 -Green 220 -Blue 120 -ExifDateTimeOriginal "2025:01:02 15:37:31"

    $payloadRefRoot = Split-Path -Path $payloadTestDir -Parent
    $payloadDbFile = Join-Path -Path $payloadRefRoot -ChildPath ".smartmediasorter_db.json"

    # Test 8a: -WhatIf run computes PayloadSha256 and persists computed hashes to database
    $payloadWhatIfStats = & $dupFinderScript -Path $payloadTestDir -ReferencePath $payloadRefRoot -DatabasePath $payloadDbFile -LogPath $logFile -WhatIf -PassThru
    Assert-True ($payloadWhatIfStats.PayloadHashesComputed -eq 5) "WhatIf run computed 5 PayloadSha256 hashes"
    Assert-True (Test-Path -LiteralPath $payloadDbFile -PathType Leaf) "WhatIf run saved computed hashes to .smartmediasorter_db.json"

    # Test 8b: Real run immediately reuses cached PayloadSha256 hashes (PayloadHashesComputed == 0)
    $payloadStats = & $dupFinderScript -Path $payloadTestDir -ReferencePath $payloadRefRoot -DatabasePath $payloadDbFile -LogPath $logFile -PassThru
    $payloadDupDir = Join-Path -Path $payloadTestDir -ChildPath "_Duplicates"

    Assert-True ($payloadStats.TotalScanned -eq 5) "Scanned 5 media files and ignored the 2 .xmp sidecar files"
    Assert-True ($payloadStats.PayloadHashesComputed -eq 0) "Real run reused all PayloadSha256 hashes cached during -WhatIf (0 recomputed)"
    Assert-True ($payloadStats.CacheHits -ge 5) "Real run recorded cache hits from -WhatIf database"
    Assert-True ($payloadStats.DuplicatesMoved -eq 2) "Detected and moved both metadata-divergent duplicates"
    Assert-True (Test-Path -LiteralPath $pairA_dash01 -PathType Leaf) "Kept 2025-01-02_153726-01.jpg (has .xmp sidecar & richer metadata)"
    Assert-True (Test-Path -LiteralPath (Join-Path -Path $payloadDupDir -ChildPath "2025-01-02_153726_000.jpg") -PathType Leaf) "Moved 2025-01-02_153726_000.jpg to _Duplicates"
    Assert-True (Test-Path -LiteralPath $pairB_plain -PathType Leaf) "Kept 2025-01-02_153731.jpg (has .xmp sidecar & richer metadata)"
    Assert-True (Test-Path -LiteralPath (Join-Path -Path $payloadDupDir -ChildPath "2025-01-02_153731_000.jpg") -PathType Leaf) "Moved 2025-01-02_153731_000.jpg to _Duplicates"
    Assert-True (Test-Path -LiteralPath $pairB_burst001 -PathType Leaf) "Preserved genuine burst shot 2025-01-02_153731_001.jpg"

    $savedPayloadDb = Get-Content -LiteralPath $payloadDbFile -Raw | ConvertFrom-Json
    Assert-True ($null -ne $savedPayloadDb.DestinationFiles.PSObject.Properties[$pairA_dash01]) "Kept file recorded in DestinationFiles (under ReferencePath)"
    Assert-True ($null -eq $savedPayloadDb.DestinationFiles.PSObject.Properties[$pairA_000]) "Moved duplicate removed from DestinationFiles"
    Assert-True ($null -ne $savedPayloadDb.DestinationDirectories.PSObject.Properties[$payloadTestDir]) "Directory under ReferencePath recorded in DestinationDirectories"

    # --- TEST 9: Stage 1C Perceptual Hash (dHash) for resized/recompressed duplicates (e.g. 2025-01-02_160119_003.jpg [larger] vs _001.jpg [smaller thumbnail + sidecars]) ---
    Write-Host "`nRunning Test 9: Stage 1C PerceptualHash (resized/recompressed duplicate 2025-01-02_160119_003.jpg vs _001.jpg)..." -ForegroundColor Cyan
    $perceptualTestDir = Join-Path -Path $testRoot -ChildPath "PicturesPerceptual\2025-01"
    New-Item -Path $perceptualTestDir -ItemType Directory -Force | Out-Null

    $pFile000_DiffPhoto = Join-Path -Path $perceptualTestDir -ChildPath "2025-01-02_160119_000.jpg"
    $pFile001_Thumb = Join-Path -Path $perceptualTestDir -ChildPath "2025-01-02_160119_001.jpg"
    $pFile001_Xmp = Join-Path -Path $perceptualTestDir -ChildPath "2025-01-02_160119_001.jpg.xmp"
    $pFile001_On1 = Join-Path -Path $perceptualTestDir -ChildPath "2025-01-02_160119_001.on1"
    $pFile003_Full = Join-Path -Path $perceptualTestDir -ChildPath "2025-01-02_160119_003.jpg"

    # Create full-resolution image (640x480) and downscaled thumbnail (160x120) of the exact same scene
    $fullBmp = New-Object System.Drawing.Bitmap(640, 480)
    $gfxFull = [System.Drawing.Graphics]::FromImage($fullBmp)
    $brush1 = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(0, 0)),
        (New-Object System.Drawing.Point(640, 480)),
        [System.Drawing.Color]::FromArgb(30, 80, 200),
        [System.Drawing.Color]::FromArgb(240, 180, 60)
    )
    $gfxFull.FillRectangle($brush1, 0, 0, 640, 480)
    $gfxFull.FillEllipse([System.Drawing.Brushes]::White, 180, 100, 280, 240)
    $gfxFull.Dispose()
    $brush1.Dispose()
    $fullBmp.Save($pFile003_Full, [System.Drawing.Imaging.ImageFormat]::Jpeg)

    $thumbBmp = New-Object System.Drawing.Bitmap(160, 120)
    $gfxThumb = [System.Drawing.Graphics]::FromImage($thumbBmp)
    $gfxThumb.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear
    $gfxThumb.DrawImage($fullBmp, 0, 0, 160, 120)
    $gfxThumb.Dispose()
    $thumbBmp.Save($pFile001_Thumb, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $thumbBmp.Dispose()
    $fullBmp.Dispose()

    [System.IO.File]::WriteAllText($pFile001_Xmp, "<x:xmpmeta>Auto sidecar on thumbnail</x:xmpmeta>")
    [System.IO.File]::WriteAllText($pFile001_On1, "{`"on1`": true}")

    # Create a genuinely different photo in the same second (_000.jpg)
    $diffBmp = New-Object System.Drawing.Bitmap(640, 480)
    $gfxDiff = [System.Drawing.Graphics]::FromImage($diffBmp)
    $gfxDiff.Clear([System.Drawing.Color]::FromArgb(20, 160, 50))
    $gfxDiff.FillRectangle([System.Drawing.Brushes]::DarkRed, 0, 0, 220, 480)
    $gfxDiff.Dispose()
    $diffBmp.Save($pFile000_DiffPhoto, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $diffBmp.Dispose()

    $pStats = & $dupFinderScript -Path $perceptualTestDir -LogPath $logFile -PassThru
    $pDupFolder = Join-Path -Path $perceptualTestDir -ChildPath "_Duplicates"

    Assert-True ($pStats.TotalScanned -eq 3) "Scanned the 3 JPEG files in 2025-01-02_160119 group"
    Assert-True ($pStats.PerceptualHashesComputed -eq 3) "Computed 3 perceptual dHash signatures"
    Assert-True ($pStats.DuplicatesMoved -eq 1) "Detected and moved the downscaled duplicate 2025-01-02_160119_001.jpg"
    Assert-True (Test-Path -LiteralPath $pFile003_Full -PathType Leaf) "Kept higher-resolution 2025-01-02_160119_003.jpg as Original despite thumbnail having sidecars"
    Assert-True (Test-Path -LiteralPath (Join-Path -Path $pDupFolder -ChildPath "2025-01-02_160119_001.jpg") -PathType Leaf) "Moved smaller 2025-01-02_160119_001.jpg to _Duplicates"
    Assert-True (Test-Path -LiteralPath (Join-Path -Path $pDupFolder -ChildPath "2025-01-02_160119_001.jpg.xmp") -PathType Leaf) "Moved associated .jpg.xmp sidecar to _Duplicates"
    Assert-True (Test-Path -LiteralPath (Join-Path -Path $pDupFolder -ChildPath "2025-01-02_160119_001.on1") -PathType Leaf) "Moved associated .on1 sidecar to _Duplicates"
    Assert-True (Test-Path -LiteralPath $pFile000_DiffPhoto -PathType Leaf) "Preserved distinct photo 2025-01-02_160119_000.jpg"

    # Regression check for .dng files containing 0xFF 0xD8 0xFF without closing 0xFF 0xD9 (must not hang in Get-ImagePerceptualHash)
    $dngSub1 = Join-Path -Path $perceptualTestDir -ChildPath "DngSub1"
    $dngSub2 = Join-Path -Path $perceptualTestDir -ChildPath "DngSub2"
    New-Item -Path $dngSub1, $dngSub2 -ItemType Directory -Force | Out-Null
    $dngA = Join-Path -Path $dngSub1 -ChildPath "2024-07-22_185127_000.dng"
    $dngB = Join-Path -Path $dngSub1 -ChildPath "2024-07-22_185127.dng"
    $dngC = Join-Path -Path $dngSub2 -ChildPath "2026-05-08_193910_000.dng"
    [byte[]]$fakeDngA = @(0x49, 0x49, 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00) + @(0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0xD9, 0x00, 0x11, 0x22, 0x33)
    [byte[]]$fakeDngB = @(0x49, 0x49, 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00) + @(0xFF, 0xD8, 0xFF, 0xDB, 0x00, 0x10, 0xD9, 0x00, 0x44, 0x55, 0x66, 0x77, 0x88)
    [System.IO.File]::WriteAllBytes($dngA, $fakeDngA)
    [System.IO.File]::WriteAllBytes($dngB, $fakeDngB)
    [System.IO.File]::WriteAllBytes($dngC, $fakeDngA)
    [System.IO.File]::WriteAllText((Join-Path -Path $dngSub1 -ChildPath ".nomedia"), "")
    [System.IO.File]::WriteAllText((Join-Path -Path $dngSub2 -ChildPath ".nomedia"), "")
    $dngStats = & $dupFinderScript -Path $perceptualTestDir -ThrottleLimit 4 -LogPath $logFile -PassThru
    Assert-True ($dngStats.Errors -eq 0 -and $dngStats.TotalScanned -eq 5) "Parallel subdirectory scan & .dng perceptual hash with unclosed 0xFF 0xD8 0xFF (and .nomedia exclusion) completed without hanging"

    # --- TEST 10: Remove-DuplicateFolders.ps1 (recursive removal of _Duplicates folders) ---
    Write-Host "`nRunning Test 10: Remove-DuplicateFolders.ps1 (WhatIf, TrashFolder mode, trash exclusion, Recycle Bin)..." -ForegroundColor Cyan
    $rmRoot = Join-Path -Path $testRoot -ChildPath "RemoveRoot"
    $rmTrash = Join-Path -Path $rmRoot -ChildPath "_DuplicatesTrash"
    $rmDupTop = Join-Path -Path $rmRoot -ChildPath "_Duplicates"
    $rmDupA = Join-Path -Path $rmRoot -ChildPath "A\_Duplicates"
    $rmDupDeep = Join-Path -Path $rmRoot -ChildPath "A\B\C\_duplicates"   # case-insensitive match
    $rmKeepFile = Join-Path -Path $rmRoot -ChildPath "A\keep.jpg"
    $rmLookAlike = Join-Path -Path $rmRoot -ChildPath "NotDuplicates\file.jpg"

    foreach ($dir in @($rmDupTop, $rmDupA, (Join-Path -Path $rmDupDeep -ChildPath "Nested"), (Split-Path -Path $rmLookAlike -Parent))) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText((Join-Path -Path $rmDupTop -ChildPath "dup1.jpg"), "dup-top")
    [System.IO.File]::WriteAllText((Join-Path -Path $rmDupA -ChildPath "dup2.jpg"), "dup-a")
    [System.IO.File]::WriteAllText((Join-Path -Path $rmDupDeep -ChildPath "Nested\dup3.jpg"), "dup-deep")
    $readOnlyDup = Join-Path -Path $rmDupDeep -ChildPath "readonly.jpg"
    [System.IO.File]::WriteAllText($readOnlyDup, "read-only duplicate")
    (Get-Item -LiteralPath $readOnlyDup).IsReadOnly = $true
    [System.IO.File]::WriteAllText($rmKeepFile, "keep me")
    [System.IO.File]::WriteAllText($rmLookAlike, "keep me too")

    # Test 10a: -WhatIf lists folders and changes nothing
    $rmWhatIf = & $removerScript -Path $rmRoot -Mode TrashFolder -TrashPath $rmTrash -LogPath $logFile -WhatIf -PassThru
    Assert-True ($rmWhatIf.FoldersFound -eq 3) "WhatIf found 3 _Duplicates folders (including case-insensitive deep match)"
    Assert-True ($rmWhatIf.FoldersRemoved -eq 0) "WhatIf removed nothing"
    Assert-True ((Test-Path -LiteralPath $rmDupTop) -and (Test-Path -LiteralPath $rmDupA) -and (Test-Path -LiteralPath $rmDupDeep)) "All _Duplicates folders still exist after -WhatIf"
    Assert-True (-not (Test-Path -LiteralPath $rmTrash)) "Trash folder was not created during -WhatIf"

    # Test 10b: TrashFolder mode moves folders into dated trash session mirroring relative paths
    $rmRun = & $removerScript -Path $rmRoot -Mode TrashFolder -TrashPath $rmTrash -LogPath $logFile -Force -PassThru
    Assert-True ($rmRun.FoldersRemoved -eq 3) "Removed 3 _Duplicates folders in TrashFolder mode"
    Assert-True ($rmRun.FilesRemoved -eq 4) "Reported 4 removed files (including nested and read-only file)"
    Assert-True ($rmRun.Errors -eq 0) "Zero errors during removal"
    Assert-True (-not (Test-Path -LiteralPath $rmDupTop) -and -not (Test-Path -LiteralPath $rmDupA) -and -not (Test-Path -LiteralPath $rmDupDeep)) "All _Duplicates folders removed from scanned tree"
    Assert-True ((Test-Path -LiteralPath $rmKeepFile -PathType Leaf) -and (Test-Path -LiteralPath $rmLookAlike -PathType Leaf)) "Regular files and look-alike folder 'NotDuplicates' preserved"
    $trashedFiles = @(Get-ChildItem -LiteralPath $rmRun.TrashSession -Recurse -File -Force | Select-Object -ExpandProperty Name)
    Assert-True ($trashedFiles.Count -eq 4 -and $trashedFiles -contains "dup3.jpg" -and $trashedFiles -contains "readonly.jpg") "All 4 files recoverable in trash session '$($rmRun.TrashSession)'"
    $rmRelative = $rmDupA.Substring([System.IO.Path]::GetPathRoot($rmDupA).Length)
    Assert-True (Test-Path -LiteralPath (Join-Path -Path $rmRun.TrashSession -ChildPath "$rmRelative\dup2.jpg") -PathType Leaf) "Trash mirrors original path relative to volume root"

    # Test 10c: Re-run finds nothing (trash folder inside scan root is excluded)
    $rmRerun = & $removerScript -Path $rmRoot -Mode TrashFolder -TrashPath $rmTrash -LogPath $logFile -Force -PassThru
    Assert-True ($rmRerun.FoldersFound -eq 0) "Re-run found 0 folders (trash folder excluded from scanning)"

    # Test 10d: Refuses to run when scan root itself is a _Duplicates folder
    $refused = $false
    try { & $removerScript -Path (Join-Path -Path $rmRun.TrashSession -ChildPath $rmRelative) -Mode TrashFolder -TrashPath $rmTrash -LogPath $logFile -Force | Out-Null }
    catch { $refused = $true }
    Assert-True $refused "Refused to remove a scan root that is itself a _Duplicates folder"

    # Test 10e: Auto mode on a local fixed drive sends folders to the Windows Recycle Bin
    $tempDriveType = ([System.IO.DriveInfo]::new([System.IO.Path]::GetPathRoot($testRoot))).DriveType
    if ($tempDriveType -eq [System.IO.DriveType]::Fixed) {
        $rbRoot = Join-Path -Path $testRoot -ChildPath "RecycleRoot"
        $rbDup = Join-Path -Path $rbRoot -ChildPath "Album\_Duplicates"
        New-Item -Path $rbDup -ItemType Directory -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path -Path $rbDup -ChildPath "SmartMediaSorter_test_dup.jpg"), "recycle bin test")
        $rbRun = & $removerScript -Path $rbRoot -LogPath $logFile -Force -PassThru
        Assert-True ($rbRun.Mode -eq "RecycleBin") "Auto mode selected RecycleBin on fixed drive"
        Assert-True ($rbRun.FoldersRemoved -eq 1 -and -not (Test-Path -LiteralPath $rbDup)) "Folder sent to Recycle Bin"
    }
    else {
        Write-Host "[SKIP] Test 10e: temp directory is not on a fixed drive" -ForegroundColor Yellow
    }

    # -------------------------------------------------------------------------
    # Test 11: Ignoring Windows Recycle Bin ($RECYCLE.BIN) and NAS Trash (#recycle)
    # -------------------------------------------------------------------------
    Write-Host "`nRunning Test 11: Ignoring Recycle Bin and Trash directories..."
    $tbRoot = Join-Path -Path $testRoot -ChildPath "TrashTestRoot"
    $tbSource = Join-Path -Path $tbRoot -ChildPath "Source"
    $tbDest = Join-Path -Path $tbRoot -ChildPath "Dest"
    $tbRecycleBin = Join-Path -Path $tbSource -ChildPath '$RECYCLE.BIN'
    $tbNasRecycle = Join-Path -Path $tbSource -ChildPath '#recycle'
    $tbNormalAlbum = Join-Path -Path $tbSource -ChildPath "NormalAlbum"

    New-Item -Path $tbRecycleBin -ItemType Directory -Force | Out-Null
    New-Item -Path $tbNasRecycle -ItemType Directory -Force | Out-Null
    New-Item -Path $tbNormalAlbum -ItemType Directory -Force | Out-Null
    New-Item -Path $tbDest -ItemType Directory -Force | Out-Null

    $normalFile = Join-Path -Path $tbNormalAlbum -ChildPath "VALID_PHOTO.jpg"
    $deletedFile1 = Join-Path -Path $tbRecycleBin -ChildPath "DELETED_PHOTO_1.jpg"
    $deletedFile2 = Join-Path -Path $tbNasRecycle -ChildPath "DELETED_PHOTO_2.jpg"

    New-SampleJpeg -FilePath $normalFile -Red 10 -Green 20 -Blue 30
    New-SampleJpeg -FilePath $deletedFile1 -Red 40 -Green 50 -Blue 60
    New-SampleJpeg -FilePath $deletedFile2 -Red 70 -Green 80 -Blue 90

    # Test 11a: SmartMediaSorter ignores $RECYCLE.BIN and #recycle
    $sortResult = & $sorterScript -SourcePath $tbSource -DestinationPath $tbDest -CopyOnly -LogPath $logFile -PassThru
    Assert-True ($sortResult.TotalScanned -eq 1) "SmartMediaSorter scanned only 1 valid file (ignored files in `$RECYCLE.BIN and #recycle)"
    Assert-True ($sortResult.Copied -eq 1) "SmartMediaSorter copied only the valid photo"
    Assert-True (Test-Path -LiteralPath $deletedFile1) "File in `$RECYCLE.BIN was untouched"
    Assert-True (Test-Path -LiteralPath $deletedFile2) "File in #recycle was untouched"

    # Test 11b: Find-MediaDuplicates ignores $RECYCLE.BIN and #recycle
    $dupResult = & $dupFinderScript -Path $tbSource -LogPath $logFile -PassThru
    Assert-True ($dupResult.TotalScanned -eq 1) "Find-MediaDuplicates scanned only 1 valid file (ignored `$RECYCLE.BIN and #recycle)"

    # Test 11c: Remove-DuplicateFolders ignores $RECYCLE.BIN and #recycle
    $trashDupFolder = Join-Path -Path $tbRecycleBin -ChildPath "_Duplicates"
    New-Item -Path $trashDupFolder -ItemType Directory -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path -Path $trashDupFolder -ChildPath "SmartMediaSorter_dummy.jpg"), "trash dummy")
    $rmTrashRun = & $removerScript -Path $tbSource -Mode TrashFolder -TrashPath (Join-Path -Path $tbRoot -ChildPath "_DuplicatesTrash") -LogPath $logFile -Force -PassThru
    Assert-True ($rmTrashRun.FoldersFound -eq 0) "Remove-DuplicateFolders did not search inside `$RECYCLE.BIN"
    Assert-True (Test-Path -LiteralPath $trashDupFolder) "Folder _Duplicates inside `$RECYCLE.BIN was untouched"

    # Test 11d: All scripts refuse to run if Path is a Recycle Bin or inside it
    $refusedSorter = $false
    try { & $sorterScript -SourcePath $tbRecycleBin -DestinationPath $tbDest -LogPath $logFile | Out-Null } catch { $refusedSorter = $true }
    Assert-True $refusedSorter "SmartMediaSorter refused to run with SourcePath set to `$RECYCLE.BIN"

    $refusedDupFinder = $false
    try { & $dupFinderScript -Path $tbRecycleBin -LogPath $logFile | Out-Null } catch { $refusedDupFinder = $true }
    Assert-True $refusedDupFinder "Find-MediaDuplicates refused to run with Path set to `$RECYCLE.BIN"

    $refusedRemover = $false
    try { & $removerScript -Path $tbRecycleBin -LogPath $logFile -Force | Out-Null } catch { $refusedRemover = $true }
    Assert-True $refusedRemover "Remove-DuplicateFolders refused to run with Path set to `$RECYCLE.BIN"

    Write-Host "`nALL TESTS PASSED SUCCESSFULLY!" -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
