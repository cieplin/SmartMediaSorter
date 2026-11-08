#Requires -Version 7.0

<#
.SYNOPSIS
    Shared helper functions for SmartMediaSorter and Find-MediaDuplicates:
    path normalization, filename stem normalization, SHA-256 whole-file and
    metadata-stripped image payload hashing, directory state hashing, sidecar
    detection, logging, and persistent JSON database management (.smartmediasorter_db.json).
#>

Set-StrictMode -Version Latest

function Write-SorterLog {
    <#
    .SYNOPSIS
        Writes a timestamped entry to the log file and optionally to the console.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet("INFO", "MOVED", "COPIED", "DELETED", "DUPLICATE", "FALLBACK", "WHATIF", "WARN", "ERROR")]
        [string]$Level = "INFO",

        [Parameter(Mandatory = $true)]
        [string]$LogFile
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timestamp] [$Level] $Message"

    Add-Content -LiteralPath $LogFile -Value $logLine -Encoding UTF8 -WhatIf:$false

    switch ($Level) {
        "ERROR"     { Write-Host $logLine -ForegroundColor Red }
        "WARN"      { Write-Host $logLine -ForegroundColor Yellow }
        "DUPLICATE" { Write-Host $logLine -ForegroundColor DarkYellow }
        "FALLBACK"  { Write-Verbose $logLine }
        "MOVED"     { Write-Host $logLine -ForegroundColor Green }
        "COPIED"    { Write-Host $logLine -ForegroundColor Cyan }
        "DELETED"   { Write-Host $logLine -ForegroundColor DarkRed }
        "WHATIF"    { Write-Host $logLine -ForegroundColor Magenta }
        default     { Write-Verbose $logLine }
    }
}

function Get-NormalizedPath {
    <#
    .SYNOPSIS
        Returns a canonical Windows full path without trailing backslash for reliable dictionary lookups.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $windowsStyle = $Path -replace '/', '\'
    $full = [System.IO.Path]::GetFullPath($windowsStyle)
    if ($full.Length -gt 3 -and $full.EndsWith('\')) {
        $full = $full.TrimEnd('\')
    }
    return $full
}

function Test-IsRecycleBinOrTrashDirectoryName {
    <#
    .SYNOPSIS
        Checks if a directory name matches Windows Recycle Bin, NAS recycle bins, or standard trash folder names.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$DirectoryName
    )

    if ([string]::IsNullOrWhiteSpace($DirectoryName)) {
        return $false
    }

    $trimmed = $DirectoryName.Trim()
    if ($trimmed.Equals('$RECYCLE.BIN', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('RECYCLER', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('RECYCLED', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('System Volume Information', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('#recycle', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('@Recycle', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('.Trash', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('.Trashes', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.Equals('_DuplicatesTrash', [System.StringComparison]::OrdinalIgnoreCase) -or
        $trimmed.StartsWith('.Trash-', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    return $false
}

function Test-IsRecycleBinOrTrashPath {
    <#
    .SYNOPSIS
        Checks if any path component in the given path belongs to a Recycle Bin or trash folder.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    try {
        $normalized = Get-NormalizedPath -Path $Path
    }
    catch {
        $normalized = $Path -replace '/', '\'
    }

    [char[]]$separators = @([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $segments = $normalized.Split($separators, [System.StringSplitOptions]::RemoveEmptyEntries)

    foreach ($seg in $segments) {
        if (Test-IsRecycleBinOrTrashDirectoryName -DirectoryName $seg) {
            return $true
        }
    }

    return $false
}

function Get-NormalizedBaseName {
    <#
    .SYNOPSIS
        Normalizes a file's base name (without extension) to group likely duplicate candidates.
        Strips common OS copy markers (" - Copy", " - kopia", " (1)", "_copy") as well as
        timestamp sequence/revision suffixes (e.g. YYYY-MM-DD_HHMMSS_000, YYYY-MM-DD_HHMMSS-01 -> YYYY-MM-DD_HHMMSS).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$BaseName
    )

    $clean = $BaseName.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $clean
    }

    $timestampSeqPattern = '^(\d{4}-\d{2}-\d{2}[_\-]\d{6})(?:[_\-]\d{1,4})+$'
    $copyPattern = '(?i)(\s*-\s*(copy|kopia)(\s*\(\d+\)|\s+\d+)?|\s*[_\-\s]+(copy|kopia)(\s*\(\d+\)|\s*_\d+|\s+\d+)?|\s*\(\d+\)|\s*\[\d+\])$'

    $changed = $true
    while ($changed) {
        $changed = $false

        if ($clean -match $copyPattern) {
            $stripped = ($clean -replace $copyPattern, '').Trim()
            if (-not [string]::IsNullOrWhiteSpace($stripped) -and $stripped -ne $clean) {
                $clean = $stripped
                $changed = $true
            }
        }

        if ($clean -match $timestampSeqPattern) {
            $strippedTs = $Matches[1]
            if ($strippedTs -ne $clean) {
                $clean = $strippedTs
                $changed = $true
            }
        }
    }

    return $clean
}

function Get-AssociatedSidecarFiles {
    <#
    .SYNOPSIS
        Returns existing sidecar file paths (.xmp, .on1, .dop, .pp3) associated with a media file.
        Checks both "<FullName>.xmp" (e.g. "photo.jpg.xmp") and "<Directory>\<BaseName>.xmp" / ".on1".
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$FileItem
    )

    $found = [System.Collections.Generic.List[string]]::new()
    $sidecarExtensions = @(".xmp", ".on1", ".dop", ".pp3")
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileItem.Name)

    foreach ($scExt in $sidecarExtensions) {
        # 1. Exact filename + sidecar extension (e.g. 2025-01-02_153726-01.jpg.xmp)
        $exactSidecar = $FileItem.FullName + $scExt
        if ([System.IO.File]::Exists($exactSidecar)) {
            $found.Add((Get-NormalizedPath -Path $exactSidecar))
        }

        # 2. Stem + sidecar extension (e.g. 2025-07-20_000749_002.on1 or 2025-01-02_153731.xmp)
        $stemSidecar = [System.IO.Path]::Combine($FileItem.DirectoryName, ($baseName + $scExt))
        if ([System.IO.File]::Exists($stemSidecar)) {
            $normStemSc = Get-NormalizedPath -Path $stemSidecar
            if (-not $found.Contains($normStemSc)) {
                $found.Add($normStemSc)
            }
        }
    }

    return ,$found
}

function Test-HasSidecarFile {
    <#
    .SYNOPSIS
        Returns $true if the specified media file has at least one associated sidecar file (.xmp, .on1, etc.).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$FileItem
    )

    $sidecars = Get-AssociatedSidecarFiles -FileItem $FileItem
    return ($sidecars.Count -gt 0)
}

function Get-StringSha256 {
    <#
    .SYNOPSIS
        Computes a SHA-256 hex digest for an arbitrary string in memory.
    #>
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$InputText)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputText)
        $hashBytes = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hashBytes) -replace '-', '')
    }
    finally {
        $sha.Dispose()
    }
}

function Read-TiffUInt16 {
    param([byte[]]$Buffer, [int]$Offset, [bool]$LittleEndian)
    if ($LittleEndian) {
        return [int][System.BitConverter]::ToUInt16($Buffer, $Offset)
    }
    return ([int]$Buffer[$Offset] -shl 8) -bor [int]$Buffer[$Offset + 1]
}

function Read-TiffUInt32 {
    param([byte[]]$Buffer, [int]$Offset, [bool]$LittleEndian)
    if ($LittleEndian) {
        return [long][System.BitConverter]::ToUInt32($Buffer, $Offset)
    }
    $b0 = [long]$Buffer[$Offset]
    $b1 = [long]$Buffer[$Offset + 1]
    $b2 = [long]$Buffer[$Offset + 2]
    $b3 = [long]$Buffer[$Offset + 3]
    return ($b0 -shl 24) -bor ($b1 -shl 16) -bor ($b2 -shl 8) -bor $b3
}

function Read-TiffUIntArray {
    param(
        [byte[]]$Buffer,
        [int]$EntryOffset,
        [int]$FieldType,
        [long]$Count,
        [bool]$LittleEndian
    )

    $result = [System.Collections.Generic.List[long]]::new()
    if ($Count -le 0 -or $Count -gt 65536) {
        return ,$result
    }

    # Type 3 = SHORT (2 bytes), Type 4 = LONG (4 bytes), Type 13 = IFD (4 bytes)
    $elemSize = if ($FieldType -eq 3) { 2 } elseif ($FieldType -eq 4 -or $FieldType -eq 13) { 4 } else { 0 }
    if ($elemSize -eq 0) {
        return ,$result
    }

    $totalBytes = $Count * $elemSize
    $dataOffset = if ($totalBytes -le 4) {
        $EntryOffset + 8
    }
    else {
        [int](Read-TiffUInt32 -Buffer $Buffer -Offset ($EntryOffset + 8) -LittleEndian $LittleEndian)
    }

    if ($dataOffset -lt 0 -or ($dataOffset + $totalBytes) -gt $Buffer.Length) {
        return ,$result
    }

    for ($i = 0; $i -lt $Count; $i++) {
        $pos = $dataOffset + ($i * $elemSize)
        $val = if ($elemSize -eq 2) {
            [long](Read-TiffUInt16 -Buffer $Buffer -Offset $pos -LittleEndian $LittleEndian)
        }
        else {
            Read-TiffUInt32 -Buffer $Buffer -Offset $pos -LittleEndian $LittleEndian
        }
        $result.Add($val)
    }

    return ,$result
}

function Get-ImagePayloadSha256 {
    <#
    .SYNOPSIS
        Computes a metadata-insensitive SHA-256 hash of the compressed/raw image stream
        for JPEG (.jpg, .jpeg), WebP (.webp), PNG (.png), and TIFF/DNG/RAW (.dng, .tif, .tiff, .cr2, .nef, .arw)
        files by skipping metadata segments (EXIF, XMP, IPTC, ICC, preview IFDs).
        Returns $null if the file is not a supported image container or cannot be parsed.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
    $tiffBasedExts = @(".dng", ".tif", ".tiff", ".cr2", ".nef", ".arw", ".orf", ".pef")
    if ($ext -notin (@(".jpg", ".jpeg", ".webp", ".png") + $tiffBasedExts)) {
        return $null
    }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -lt 16) {
            return $null
        }

        $sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            # --- 1. JPEG (.jpg, .jpeg) ---
            if (($ext -eq ".jpg" -or $ext -eq ".jpeg") -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8) {
                $pos = 2
                $hashedAnyImageSegment = $false

                while ($pos + 2 -le $bytes.Length) {
                    if ($bytes[$pos] -ne 0xFF) {
                        break
                    }

                    # Skip padding 0xFF bytes
                    while ($pos -lt $bytes.Length -and $bytes[$pos] -eq 0xFF) {
                        $pos++
                    }
                    if ($pos -ge $bytes.Length) {
                        break
                    }

                    $marker = $bytes[$pos]
                    $pos++

                    # EOI (0xD9) or SOI (0xD8) or TEM (0x01) or RST0..RST7 (0xD0..0xD7) have no length field
                    if ($marker -eq 0xD9) {
                        break
                    }
                    if ($marker -eq 0xD8 -or $marker -eq 0x01 -or ($marker -ge 0xD0 -and $marker -le 0xD7)) {
                        continue
                    }

                    if ($pos + 2 -gt $bytes.Length) {
                        break
                    }

                    $segLen = ([int]$bytes[$pos] -shl 8) -bor [int]$bytes[$pos + 1]
                    if ($segLen -lt 2 -or ($pos + $segLen) -gt $bytes.Length) {
                        break
                    }

                    # SOS (0xDA): Start of Scan -> hash SOS header + entropy-coded bitstream up to EOI (0xFF 0xD9)
                    if ($marker -eq 0xDA) {
                        $scanDataStart = $pos + $segLen
                        # Fast EOI (0xFF 0xD9) lookup: check end of file first, otherwise use native SIMD [Array]::IndexOf
                        $scanEnd = $bytes.Length
                        if ($bytes.Length -ge ($scanDataStart + 2) -and
                            $bytes[$bytes.Length - 2] -eq 0xFF -and $bytes[$bytes.Length - 1] -eq 0xD9) {
                            $scanEnd = $bytes.Length - 2
                        }
                        else {
                            $searchIdx = $scanDataStart + 1
                            while ($searchIdx -lt $bytes.Length) {
                                $d9Pos = [Array]::IndexOf($bytes, [byte]0xD9, $searchIdx)
                                if ($d9Pos -lt 0) {
                                    break
                                }
                                if ($bytes[$d9Pos - 1] -eq 0xFF) {
                                    $scanEnd = $d9Pos - 1
                                    break
                                }
                                $searchIdx = $d9Pos + 1
                            }
                        }

                        # Hash SOS header and entropy scan stream
                        $totalScanBytes = $scanEnd - ($pos - 2)
                        if ($totalScanBytes -gt 0) {
                            [void]$sha.TransformBlock($bytes, $pos - 2, $totalScanBytes, $null, 0)
                            $hashedAnyImageSegment = $true
                        }
                        break
                    }

                    # Skip APP0..APP15 (0xE0..0xEF) and COM (0xFE) metadata segments
                    $isMetadataSegment = (($marker -ge 0xE0 -and $marker -le 0xEF) -or ($marker -eq 0xFE))
                    if (-not $isMetadataSegment) {
                        # Keep image tables (DQT 0xDB, DHT 0xC4, SOF 0xC0..0xCF, DRI 0xDD)
                        [void]$sha.TransformBlock($bytes, $pos - 2, $segLen + 2, $null, 0)
                        $hashedAnyImageSegment = $true
                    }

                    $pos += $segLen
                }

                if ($hashedAnyImageSegment) {
                    [void]$sha.TransformFinalBlock($bytes, 0, 0)
                    return ([System.BitConverter]::ToString($sha.Hash) -replace '-', '')
                }
                return $null
            }

            # --- 2. WebP (.webp) ---
            if ($ext -eq ".webp" -and $bytes.Length -ge 20 -and
                $bytes[0] -eq 0x52 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46 -and $bytes[3] -eq 0x46 -and
                $bytes[8] -eq 0x57 -and $bytes[9] -eq 0x45 -and $bytes[10] -eq 0x42 -and $bytes[11] -eq 0x50) {

                $pos = 12
                $hashedAnyChunk = $false
                while ($pos + 8 -le $bytes.Length) {
                    $chunkFourCC = [System.Text.Encoding]::ASCII.GetString($bytes, $pos, 4)
                    $chunkSize = [System.BitConverter]::ToInt32($bytes, $pos + 4)
                    if ($chunkSize -lt 0 -or ($pos + 8 + $chunkSize) -gt $bytes.Length) {
                        break
                    }

                    # Hash only visual image data chunks (VP8, VP8L, ALPH, ANMF), skipping EXIF, XMP, ICCP
                    if ($chunkFourCC -in @("VP8 ", "VP8L", "ALPH", "ANMF")) {
                        [void]$sha.TransformBlock($bytes, $pos, 8 + $chunkSize, $null, 0)
                        $hashedAnyChunk = $true
                    }

                    $paddedSize = $chunkSize + ($chunkSize % 2)
                    $pos += 8 + $paddedSize
                }

                if ($hashedAnyChunk) {
                    [void]$sha.TransformFinalBlock($bytes, 0, 0)
                    return ([System.BitConverter]::ToString($sha.Hash) -replace '-', '')
                }
                return $null
            }

            # --- 3. PNG (.png) ---
            if ($ext -eq ".png" -and $bytes.Length -ge 24 -and
                $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) {

                $pos = 8
                $hashedAnyChunk = $false
                while ($pos + 12 -le $bytes.Length) {
                    $chunkLen = ([int]$bytes[$pos] -shl 24) -bor ([int]$bytes[$pos + 1] -shl 16) -bor ([int]$bytes[$pos + 2] -shl 8) -bor [int]$bytes[$pos + 3]
                    $chunkType = [System.Text.Encoding]::ASCII.GetString($bytes, $pos + 4, 4)
                    if ($chunkLen -lt 0 -or ($pos + 12 + $chunkLen) -gt $bytes.Length) {
                        break
                    }

                    if ($chunkType -in @("IHDR", "PLTE", "tRNS", "IDAT")) {
                        [void]$sha.TransformBlock($bytes, $pos + 4, 4 + $chunkLen, $null, 0)
                        $hashedAnyChunk = $true
                    }
                    if ($chunkType -eq "IEND") {
                        break
                    }

                    $pos += 12 + $chunkLen
                }

                if ($hashedAnyChunk) {
                    [void]$sha.TransformFinalBlock($bytes, 0, 0)
                    return ([System.BitConverter]::ToString($sha.Hash) -replace '-', '')
                }
                return $null
            }

            # --- 4. TIFF / DNG / RAW (.dng, .tif, .tiff, .cr2, .nef, .arw, .orf, .pef) ---
            if ($ext -in $tiffBasedExts -and $bytes.Length -ge 16 -and
                (($bytes[0] -eq 0x49 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x2A -and $bytes[3] -eq 0x00) -or
                 ($bytes[0] -eq 0x4D -and $bytes[1] -eq 0x4D -and $bytes[2] -eq 0x00 -and $bytes[3] -eq 0x2A))) {

                $isLe = ($bytes[0] -eq 0x49)
                $ifd0Offset = [int](Read-TiffUInt32 -Buffer $bytes -Offset 4 -LittleEndian $isLe)

                $ifdQueue = [System.Collections.Generic.Queue[int]]::new()
                $visitedIfds = [System.Collections.Generic.HashSet[int]]::new()
                if ($ifd0Offset -ge 8 -and $ifd0Offset -lt $bytes.Length) {
                    $ifdQueue.Enqueue($ifd0Offset)
                }

                $bestOffsets = $null
                $bestByteCounts = $null
                $bestTotalBytes = 0L

                while ($ifdQueue.Count -gt 0) {
                    $ifdOff = $ifdQueue.Dequeue()
                    if (-not $visitedIfds.Add($ifdOff)) {
                        continue
                    }
                    if ($ifdOff -lt 8 -or ($ifdOff + 2) -gt $bytes.Length) {
                        continue
                    }

                    $entryCount = Read-TiffUInt16 -Buffer $bytes -Offset $ifdOff -LittleEndian $isLe
                    if ($entryCount -le 0 -or $entryCount -gt 512) {
                        continue
                    }

                    $stripOffsets = $null
                    $stripCounts = $null
                    $tileOffsets = $null
                    $tileCounts = $null

                    for ($i = 0; $i -lt $entryCount; $i++) {
                        $ePos = $ifdOff + 2 + ($i * 12)
                        if (($ePos + 12) -gt $bytes.Length) {
                            break
                        }

                        $tag = Read-TiffUInt16 -Buffer $bytes -Offset $ePos -LittleEndian $isLe
                        $fType = Read-TiffUInt16 -Buffer $bytes -Offset ($ePos + 2) -LittleEndian $isLe
                        $vCount = Read-TiffUInt32 -Buffer $bytes -Offset ($ePos + 4) -LittleEndian $isLe

                        # 0x014A = SubIFDs
                        if ($tag -eq 0x014A) {
                            $subOffsets = Read-TiffUIntArray -Buffer $bytes -EntryOffset $ePos -FieldType $fType -Count $vCount -LittleEndian $isLe
                            foreach ($so in $subOffsets) {
                                if ($so -ge 8 -and $so -lt $bytes.Length) {
                                    $ifdQueue.Enqueue([int]$so)
                                }
                            }
                        }
                        # 0x0111 = StripOffsets
                        elseif ($tag -eq 0x0111) {
                            $stripOffsets = Read-TiffUIntArray -Buffer $bytes -EntryOffset $ePos -FieldType $fType -Count $vCount -LittleEndian $isLe
                        }
                        # 0x0117 = StripByteCounts
                        elseif ($tag -eq 0x0117) {
                            $stripCounts = Read-TiffUIntArray -Buffer $bytes -EntryOffset $ePos -FieldType $fType -Count $vCount -LittleEndian $isLe
                        }
                        # 0x0144 = TileOffsets
                        elseif ($tag -eq 0x0144) {
                            $tileOffsets = Read-TiffUIntArray -Buffer $bytes -EntryOffset $ePos -FieldType $fType -Count $vCount -LittleEndian $isLe
                        }
                        # 0x0145 = TileByteCounts
                        elseif ($tag -eq 0x0145) {
                            $tileCounts = Read-TiffUIntArray -Buffer $bytes -EntryOffset $ePos -FieldType $fType -Count $vCount -LittleEndian $isLe
                        }
                    }

                    # Check next IFD pointer after entries
                    $nextIfdPos = $ifdOff + 2 + ($entryCount * 12)
                    if (($nextIfdPos + 4) -le $bytes.Length) {
                        $nextIfdOff = [int](Read-TiffUInt32 -Buffer $bytes -Offset $nextIfdPos -LittleEndian $isLe)
                        if ($nextIfdOff -ge 8 -and $nextIfdOff -lt $bytes.Length) {
                            $ifdQueue.Enqueue($nextIfdOff)
                        }
                    }

                    foreach ($pair in @(
                        @{ Offsets = $stripOffsets; Counts = $stripCounts },
                        @{ Offsets = $tileOffsets;  Counts = $tileCounts }
                    )) {
                        $offs = $pair.Offsets
                        $cnts = $pair.Counts
                        if ($null -ne $offs -and $null -ne $cnts -and $offs.Count -gt 0 -and $offs.Count -eq $cnts.Count) {
                            $sumBytes = 0L
                            $validBlocks = $true
                            for ($j = 0; $j -lt $offs.Count; $j++) {
                                if ($offs[$j] -lt 8 -or $cnts[$j] -le 0 -or ($offs[$j] + $cnts[$j]) -gt $bytes.Length) {
                                    $validBlocks = $false
                                    break
                                }
                                $sumBytes += $cnts[$j]
                            }
                            if ($validBlocks -and $sumBytes -gt $bestTotalBytes) {
                                $bestTotalBytes = $sumBytes
                                $bestOffsets = $offs
                                $bestByteCounts = $cnts
                            }
                        }
                    }
                }

                if ($bestTotalBytes -gt 0 -and $null -ne $bestOffsets) {
                    for ($k = 0; $k -lt $bestOffsets.Count; $k++) {
                        [void]$sha.TransformBlock($bytes, [int]$bestOffsets[$k], [int]$bestByteCounts[$k], $null, 0)
                    }
                    [void]$sha.TransformFinalBlock($bytes, 0, 0)
                    return ([System.BitConverter]::ToString($sha.Hash) -replace '-', '')
                }
                return $null
            }
        }
        finally {
            $sha.Dispose()
        }
    }
    catch {
        return $null
    }

    return $null
}

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

function Get-BitmapPerceptualSignature {
    <#
    .SYNOPSIS
        Computes a compact perceptual signature "<Orient>:<RGB3>:<dHash16>" from a System.Drawing.Image
        using a 9x8 grayscale luminance grid (64-bit horizontal gradient dHash) plus aspect orientation
        (L/P/S) and 12-bit quantized average RGB color.
    #>
    param(
        [Parameter(Mandatory = $true)][System.Drawing.Image]$Image
    )

    $w = $Image.Width
    $h = $Image.Height
    if ($w -le 0 -or $h -le 0) {
        return $null
    }

    $maxDim = [double][Math]::Max($w, $h)
    $diffRatio = [Math]::Abs($w - $h) / $maxDim
    $orient = if ($diffRatio -le 0.05) { "S" } elseif ($w -gt $h) { "L" } else { "P" }

    $smallBmp = [System.Drawing.Bitmap]::new(9, 8, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    try {
        $gfx = [System.Drawing.Graphics]::FromImage($smallBmp)
        try {
            $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear
            $gfx.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighSpeed
            $gfx.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $gfx.DrawImage($Image, 0, 0, 9, 8)
        }
        finally {
            $gfx.Dispose()
        }

        $lumGrid = [int[]]::new(72)
        [long]$sumR = 0
        [long]$sumG = 0
        [long]$sumB = 0

        for ($y = 0; $y -lt 8; $y++) {
            for ($x = 0; $x -lt 9; $x++) {
                $px = $smallBmp.GetPixel($x, $y)
                $r = [int]$px.R
                $g = [int]$px.G
                $b = [int]$px.B
                $sumR += $r
                $sumG += $g
                $sumB += $b
                $lumGrid[$y * 9 + $x] = (299 * $r + 587 * $g + 114 * $b)
            }
        }

        $avgR4 = [Math]::Min(15, [Math]::Max(0, [int](($sumR / 72) -shr 4)))
        $avgG4 = [Math]::Min(15, [Math]::Max(0, [int](($sumG / 72) -shr 4)))
        $avgB4 = [Math]::Min(15, [Math]::Max(0, [int](($sumB / 72) -shr 4)))

        [uint64]$dHash = 0UL
        [uint64]$oneBit = 1UL
        $bitIdx = 0

        for ($y = 0; $y -lt 8; $y++) {
            $rowOffset = $y * 9
            for ($x = 0; $x -lt 8; $x++) {
                if ($lumGrid[$rowOffset + $x] -gt $lumGrid[$rowOffset + $x + 1]) {
                    $dHash = $dHash -bor ($oneBit -shl $bitIdx)
                }
                $bitIdx++
            }
        }

        return ("{0}:{1:X1}{2:X1}{3:X1}:{4:X16}" -f $orient, $avgR4, $avgG4, $avgB4, $dHash)
    }
    finally {
        $smallBmp.Dispose()
    }
}

function Test-IsStandardJpegStream {
    <#
    .SYNOPSIS
        Validates that a byte range is a standard 8-bit baseline/progressive JPEG (1-channel grayscale
        or 3-channel YCbCr with component IDs 1,2,3) so GDI+/libjpeg never prints
        "Unsupported color conversion request" on RAW tiles or CMYK/YCCK/lossless JPEGs.
    #>
    param(
        [Parameter(Mandatory = $true)][byte[]]$Buffer,
        [Parameter(Mandatory = $true)][int]$Start,
        [Parameter(Mandatory = $true)][int]$Length
    )

    $end = $Start + $Length
    if ($Start -lt 0 -or $Length -lt 32 -or $end -gt $Buffer.Length) {
        return $false
    }
    if ($Buffer[$Start] -ne 0xFF -or $Buffer[$Start + 1] -ne 0xD8) {
        return $false
    }

    $pos = $Start + 2
    $foundValidSof = $false

    while (($pos + 4) -le $end) {
        if ($Buffer[$pos] -ne 0xFF) {
            return $false
        }
        while ($pos -lt $end -and $Buffer[$pos] -eq 0xFF) {
            $pos++
        }
        if ($pos -ge $end) {
            return $false
        }

        $marker = $Buffer[$pos]
        $pos++

        if ($marker -eq 0xD9) {
            return $false
        }
        if ($marker -eq 0xD8 -or $marker -eq 0x01 -or ($marker -ge 0xD0 -and $marker -le 0xD7)) {
            continue
        }

        if (($pos + 2) -gt $end) {
            return $false
        }
        $segLen = ([int]$Buffer[$pos] -shl 8) -bor [int]$Buffer[$pos + 1]
        if ($segLen -lt 2 -or ($pos + $segLen) -gt $end) {
            return $false
        }

        # Check SOF markers (0xC0..0xCF except DHT=0xC4, JPG=0xC8, DAC=0xCC)
        if ($marker -ge 0xC0 -and $marker -le 0xCF -and $marker -ne 0xC4 -and $marker -ne 0xC8 -and $marker -ne 0xCC) {
            if ($marker -ne 0xC0 -and $marker -ne 0xC2) {
                return $false
            }
            if ($segLen -lt 11) {
                return $false
            }
            $precision = $Buffer[$pos + 2]
            $height = ([int]$Buffer[$pos + 3] -shl 8) -bor [int]$Buffer[$pos + 4]
            $width  = ([int]$Buffer[$pos + 5] -shl 8) -bor [int]$Buffer[$pos + 6]
            $nComp  = $Buffer[$pos + 7]

            if ($precision -ne 8 -or $height -lt 16 -or $width -lt 16) {
                return $false
            }
            if ($nComp -eq 1) {
                $foundValidSof = $true
            }
            elseif ($nComp -eq 3 -and $segLen -ge 17) {
                $c1 = $Buffer[$pos + 8]
                $c2 = $Buffer[$pos + 11]
                $c3 = $Buffer[$pos + 14]
                if ($c1 -eq 1 -and $c2 -eq 2 -and $c3 -eq 3) {
                    $foundValidSof = $true
                }
                else {
                    return $false
                }
            }
            else {
                return $false
            }
        }
        elseif ($marker -eq 0xEE -and $segLen -ge 14) {
            # APP14 Adobe marker: only allow standard YCbCr transform (1)
            if ($Buffer[$pos + 2] -eq 0x41 -and $Buffer[$pos + 3] -eq 0x64 -and $Buffer[$pos + 4] -eq 0x6F -and $Buffer[$pos + 5] -eq 0x62 -and $Buffer[$pos + 6] -eq 0x65) {
                if ($Buffer[$pos + 13] -ne 1) {
                    return $false
                }
            }
        }
        elseif ($marker -eq 0xDA) {
            if (-not $foundValidSof) {
                return $false
            }
            $sosComp = $Buffer[$pos + 2]
            if ($sosComp -ne 1 -and $sosComp -ne 3) {
                return $false
            }
            return $true
        }

        $pos += $segLen
    }

    return $false
}

function Invoke-SafeGdiPerceptualSignature {
    <#
    .SYNOPSIS
        Decodes an image from an in-memory byte slice and computes its perceptual signature
        under a process-wide Monitor lock on [System.Drawing.Bitmap] so concurrent Runspace
        threads never deadlock inside native gdiplus.dll.
    #>
    param(
        [Parameter(Mandatory = $true)][byte[]]$Buffer,
        [Parameter(Mandatory = $true)][int]$Start,
        [Parameter(Mandatory = $true)][int]$Length
    )

    $lockTarget = [System.Drawing.Bitmap]
    $lockTaken = $false
    try {
        [System.Threading.Monitor]::TryEnter($lockTarget, 30000, [ref]$lockTaken)
        if (-not $lockTaken) {
            return $null
        }

        $ms = [System.IO.MemoryStream]::new($Buffer, $Start, $Length, $false)
        try {
            $img = [System.Drawing.Image]::FromStream($ms, $false, $false)
            try {
                return (Get-BitmapPerceptualSignature -Image $img)
            }
            finally {
                $img.Dispose()
            }
        }
        finally {
            $ms.Dispose()
        }
    }
    catch {
        return $null
    }
    finally {
        if ($lockTaken) {
            [System.Threading.Monitor]::Exit($lockTarget)
        }
    }
}

function Get-ImagePerceptualHash {
    <#
    .SYNOPSIS
        Computes a resolution- and compression-insensitive perceptual hash ("<Orient>:<RGB3>:<dHash16>")
        for standard image files (.jpg, .jpeg, .png, .bmp, .gif) and RAW/DNG/TIFF files
        (by extracting the largest valid embedded JPEG preview).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath
    )

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
    $directExts = @(".jpg", ".jpeg", ".png", ".bmp", ".gif")
    $rawPreviewExts = @(".dng", ".cr2", ".nef", ".arw", ".orf", ".pef", ".tif", ".tiff")

    if ($ext -notin ($directExts + $rawPreviewExts)) {
        return $null
    }

    try {
        if ($ext -in $directExts) {
            $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
            if ($ext -eq ".jpg" -or $ext -eq ".jpeg") {
                if (-not (Test-IsStandardJpegStream -Buffer $fileBytes -Start 0 -Length $fileBytes.Length)) {
                    return $null
                }
            }
            return (Invoke-SafeGdiPerceptualSignature -Buffer $fileBytes -Start 0 -Length $fileBytes.Length)
        }
        else {
            # Extract largest valid embedded JPEG preview (0xFF 0xD8 ... 0xFF 0xD9) from RAW/DNG/TIFF
            $bytes = [System.IO.File]::ReadAllBytes($FilePath)
            $candidates = [System.Collections.Generic.List[object]]::new()
            $searchIdx = 0

            while ($searchIdx -lt ($bytes.Length - 4)) {
                $ffPos = [Array]::IndexOf($bytes, [byte]0xFF, $searchIdx)
                if ($ffPos -lt 0 -or ($ffPos + 3) -ge $bytes.Length) {
                    break
                }

                # Always advance searchIdx so outer loop strictly progresses
                $searchIdx = $ffPos + 1

                if ($bytes[$ffPos + 1] -eq 0xD8 -and $bytes[$ffPos + 2] -eq 0xFF -and $bytes[$ffPos + 3] -ge 0xC0 -and $bytes[$ffPos + 3] -le 0xFE) {
                    $searchIdx = $ffPos + 2
                    $eoiSearch = $ffPos + 4
                    $noMoreD9 = $false
                    while ($eoiSearch -lt $bytes.Length) {
                        $d9Pos = [Array]::IndexOf($bytes, [byte]0xD9, $eoiSearch)
                        if ($d9Pos -lt 0) {
                            $noMoreD9 = $true
                            break
                        }
                        if ($bytes[$d9Pos - 1] -eq 0xFF) {
                            $candLen = ($d9Pos - $ffPos) + 1
                            if ($candLen -ge 2048 -and (Test-IsStandardJpegStream -Buffer $bytes -Start $ffPos -Length $candLen)) {
                                $candidates.Add([PSCustomObject]@{
                                    Start  = $ffPos
                                    Length = $candLen
                                })
                            }
                            $searchIdx = $d9Pos + 1
                            break
                        }
                        $eoiSearch = $d9Pos + 1
                    }
                    if ($noMoreD9 -or $eoiSearch -ge $bytes.Length) {
                        break
                    }
                }
            }

            if ($candidates.Count -gt 0) {
                $sortedCandidates = @($candidates | Sort-Object -Property Length -Descending | Select-Object -First 6)
                foreach ($cand in $sortedCandidates) {
                    $sig = Invoke-SafeGdiPerceptualSignature -Buffer $bytes -Start $cand.Start -Length $cand.Length
                    if (-not [string]::IsNullOrWhiteSpace($sig)) {
                        return $sig
                    }
                }
            }
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-PerceptualHashDistance {
    <#
    .SYNOPSIS
        Computes the Hamming distance (0..64) between two perceptual signatures ("<Orient>:<RGB3>:<dHash16>").
        Returns 64 (maximum distance / no match) if orientation differs or average RGB channels differ by > 2.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$HashA,
        [Parameter(Mandatory = $true)][string]$HashB
    )

    if ([string]::IsNullOrWhiteSpace($HashA) -or [string]::IsNullOrWhiteSpace($HashB)) {
        return 64
    }
    if ($HashA.Equals($HashB, [System.StringComparison]::OrdinalIgnoreCase)) {
        return 0
    }

    $partsA = $HashA -split ':'
    $partsB = $HashB -split ':'
    if ($partsA.Count -ne 3 -or $partsB.Count -ne 3) {
        return 64
    }

    # Orientation must match (L, P, or S)
    if ($partsA[0] -ne $partsB[0]) {
        return 64
    }

    # Check 12-bit average RGB channels (each nibble 0..15; allow up to +-2 variation for recompression/color profile)
    $rgbA = $partsA[1]
    $rgbB = $partsB[1]
    if ($rgbA.Length -eq 3 -and $rgbB.Length -eq 3) {
        $rDiff = [Math]::Abs([Convert]::ToInt32($rgbA.Substring(0, 1), 16) - [Convert]::ToInt32($rgbB.Substring(0, 1), 16))
        $gDiff = [Math]::Abs([Convert]::ToInt32($rgbA.Substring(1, 1), 16) - [Convert]::ToInt32($rgbB.Substring(1, 1), 16))
        $bDiff = [Math]::Abs([Convert]::ToInt32($rgbA.Substring(2, 1), 16) - [Convert]::ToInt32($rgbB.Substring(2, 1), 16))
        if ($rDiff -gt 2 -or $gDiff -gt 2 -or $bDiff -gt 2) {
            return 64
        }
    }

    try {
        [uint64]$valA = [Convert]::ToUInt64($partsA[2], 16)
        [uint64]$valB = [Convert]::ToUInt64($partsB[2], 16)
        [uint64]$xor = $valA -bxor $valB
        $dist = 0
        while ($xor -ne 0UL) {
            $dist++
            $xor = $xor -band ($xor - 1UL)
        }
        return $dist
    }
    catch {
        return 64
    }
}

function Set-DatabaseFileRecord {
    <#
    .SYNOPSIS
        Inserts or updates a file record in the persistent database (DestinationFiles if under
        ReferenceRoot or already in DestinationFiles, otherwise SourceFiles), preserving existing
        cached hashes (Sha256, PayloadSha256, PerceptualHash) and EXIF metadata when Length and LastWriteTimeUtcTicks match.
    #>
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Database,
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$FileItem,
        [Parameter()][string]$ReferenceRoot = "",
        [Parameter()][string]$Sha256 = $null,
        [Parameter()][string]$PayloadSha256 = $null,
        [Parameter()][string]$PerceptualHash = $null
    )

    $normFile = Get-NormalizedPath -Path $FileItem.FullName
    $normDir = Get-NormalizedPath -Path $FileItem.DirectoryName

    $isDest = $Database["DestinationFiles"].ContainsKey($normFile)
    if (-not $isDest -and -not [string]::IsNullOrWhiteSpace($ReferenceRoot)) {
        $normRef = Get-NormalizedPath -Path $ReferenceRoot
        if ($normFile.StartsWith($normRef + "\", [System.StringComparison]::OrdinalIgnoreCase) -or
            $normDir.Equals($normRef, [System.StringComparison]::OrdinalIgnoreCase)) {
            $isDest = $true
        }
    }

    $targetTableName = if ($isDest) { "DestinationFiles" } else { "SourceFiles" }
    $otherTableName = if ($isDest) { "SourceFiles" } else { "DestinationFiles" }
    $targetTable = $Database[$targetTableName]
    $otherTable = $Database[$otherTableName]

    $existingRec = $null
    if ($targetTable.ContainsKey($normFile)) {
        $existingRec = $targetTable[$normFile]
    }
    elseif ($otherTable.ContainsKey($normFile)) {
        $existingRec = $otherTable[$normFile]
        $otherTable.Remove($normFile)
    }

    $newRec = @{
        Name                  = $FileItem.Name
        DirectoryPath         = $normDir
        Length                = [long]$FileItem.Length
        LastWriteTimeUtcTicks = [long]$FileItem.LastWriteTimeUtc.Ticks
        Sha256                = $null
        PayloadSha256         = $null
        PerceptualHash        = $null
    }

    if ($null -ne $existingRec) {
        $sameState = (
            $existingRec.ContainsKey("Length") -and ($existingRec["Length"] -eq $FileItem.Length) -and
            $existingRec.ContainsKey("LastWriteTimeUtcTicks") -and ($existingRec["LastWriteTimeUtcTicks"] -eq $FileItem.LastWriteTimeUtc.Ticks)
        )
        if ($sameState) {
            foreach ($preservedKey in @("Sha256", "PayloadSha256", "PerceptualHash", "PhotoDateIso", "DateSource", "LastAction", "TargetPath")) {
                if ($existingRec.ContainsKey($preservedKey) -and $null -ne $existingRec[$preservedKey]) {
                    $newRec[$preservedKey] = $existingRec[$preservedKey]
                }
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Sha256)) {
        $newRec["Sha256"] = $Sha256
    }
    if (-not [string]::IsNullOrWhiteSpace($PayloadSha256)) {
        $newRec["PayloadSha256"] = $PayloadSha256
    }
    if (-not [string]::IsNullOrWhiteSpace($PerceptualHash)) {
        $newRec["PerceptualHash"] = $PerceptualHash
    }

    $targetTable[$normFile] = $newRec
    return $newRec
}

function Get-CachedImagePerceptualHash {
    <#
    .SYNOPSIS
        Retrieves or computes the perceptual hash ("<Orient>:<RGB3>:<dHash16>") for an image file,
        caching the result in the session dictionary and in the persistent database (PerceptualHash).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][hashtable]$PerceptualCache,
        [Parameter()][System.Collections.IDictionary]$Database = $null,
        [Parameter()][System.Collections.IDictionary]$Stats = $null,
        [Parameter()][System.IO.FileInfo]$FileItem = $null,
        [Parameter()][string]$ReferenceRoot = ""
    )

    $norm = Get-NormalizedPath -Path $FilePath
    if ($PerceptualCache.ContainsKey($norm)) {
        return $PerceptualCache[$norm]
    }

    $preKey = "__precomputed__:$norm"
    if ($PerceptualCache.ContainsKey($preKey) -and -not [string]::IsNullOrWhiteSpace($PerceptualCache[$preKey])) {
        $computedPHash = [string]$PerceptualCache[$preKey]
        $PerceptualCache[$norm] = $computedPHash
        if ($null -ne $Database -and $null -ne $FileItem) {
            [void](Set-DatabaseFileRecord -Database $Database -FileItem $FileItem -ReferenceRoot $ReferenceRoot -PerceptualHash $computedPHash)
        }
        return $computedPHash
    }

    if ($null -ne $Database) {
        foreach ($tableName in @("DestinationFiles", "SourceFiles")) {
            if ($Database[$tableName].ContainsKey($norm)) {
                $rec = $Database[$tableName][$norm]
                if ($rec.ContainsKey("PerceptualHash") -and -not [string]::IsNullOrWhiteSpace($rec["PerceptualHash"])) {
                    $isMatch = $true
                    if ($null -ne $FileItem -and $rec.ContainsKey("Length") -and $rec.ContainsKey("LastWriteTimeUtcTicks")) {
                        if ($rec["Length"] -ne $FileItem.Length -or $rec["LastWriteTimeUtcTicks"] -ne $FileItem.LastWriteTimeUtc.Ticks) {
                            $isMatch = $false
                        }
                    }
                    if ($isMatch) {
                        $PerceptualCache[$norm] = [string]$rec["PerceptualHash"]
                        if ($null -ne $Stats) { $Stats.CacheHits++ }
                        return $PerceptualCache[$norm]
                    }
                }
            }
        }
    }

    $computedPHash = Get-ImagePerceptualHash -FilePath $norm
    $PerceptualCache[$norm] = $computedPHash

    if (-not [string]::IsNullOrWhiteSpace($computedPHash) -and $null -ne $Database) {
        if ($null -ne $FileItem) {
            [void](Set-DatabaseFileRecord -Database $Database -FileItem $FileItem -ReferenceRoot $ReferenceRoot -PerceptualHash $computedPHash)
        }
        else {
            if ($Database["DestinationFiles"].ContainsKey($norm)) {
                $Database["DestinationFiles"][$norm]["PerceptualHash"] = $computedPHash
            }
            if ($Database["SourceFiles"].ContainsKey($norm)) {
                $Database["SourceFiles"][$norm]["PerceptualHash"] = $computedPHash
            }
        }
    }

    return $computedPHash
}

function Get-CachedImagePayloadHash {
    <#
    .SYNOPSIS
        Retrieves or computes the metadata-stripped image payload SHA-256 hash for an image file,
        caching the result in the session dictionary and in the persistent database (PayloadSha256).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][hashtable]$PayloadCache,
        [Parameter()][System.Collections.IDictionary]$Database = $null,
        [Parameter()][System.Collections.IDictionary]$Stats = $null,
        [Parameter()][System.IO.FileInfo]$FileItem = $null,
        [Parameter()][string]$ReferenceRoot = ""
    )

    $norm = Get-NormalizedPath -Path $FilePath
    if ($PayloadCache.ContainsKey($norm)) {
        return $PayloadCache[$norm]
    }

    $preKey = "__precomputed__:$norm"
    if ($PayloadCache.ContainsKey($preKey) -and -not [string]::IsNullOrWhiteSpace($PayloadCache[$preKey])) {
        $computedPayloadHash = [string]$PayloadCache[$preKey]
        $PayloadCache[$norm] = $computedPayloadHash
        if ($null -ne $Database -and $null -ne $FileItem) {
            [void](Set-DatabaseFileRecord -Database $Database -FileItem $FileItem -ReferenceRoot $ReferenceRoot -PayloadSha256 $computedPayloadHash)
        }
        return $computedPayloadHash
    }

    if ($null -ne $Database) {
        foreach ($tableName in @("DestinationFiles", "SourceFiles")) {
            if ($Database[$tableName].ContainsKey($norm)) {
                $rec = $Database[$tableName][$norm]
                if ($rec.ContainsKey("PayloadSha256") -and -not [string]::IsNullOrWhiteSpace($rec["PayloadSha256"])) {
                    $isMatch = $true
                    if ($null -ne $FileItem -and $rec.ContainsKey("Length") -and $rec.ContainsKey("LastWriteTimeUtcTicks")) {
                        if ($rec["Length"] -ne $FileItem.Length -or $rec["LastWriteTimeUtcTicks"] -ne $FileItem.LastWriteTimeUtc.Ticks) {
                            $isMatch = $false
                        }
                    }
                    if ($isMatch) {
                        $PayloadCache[$norm] = [string]$rec["PayloadSha256"]
                        if ($null -ne $Stats) { $Stats.CacheHits++ }
                        return $PayloadCache[$norm]
                    }
                }
            }
        }
    }

    $computedPayloadHash = Get-ImagePayloadSha256 -FilePath $norm
    $PayloadCache[$norm] = $computedPayloadHash

    if (-not [string]::IsNullOrWhiteSpace($computedPayloadHash) -and $null -ne $Database) {
        if ($null -ne $FileItem) {
            [void](Set-DatabaseFileRecord -Database $Database -FileItem $FileItem -ReferenceRoot $ReferenceRoot -PayloadSha256 $computedPayloadHash)
        }
        else {
            if ($Database["DestinationFiles"].ContainsKey($norm)) {
                $Database["DestinationFiles"][$norm]["PayloadSha256"] = $computedPayloadHash
            }
            if ($Database["SourceFiles"].ContainsKey($norm)) {
                $Database["SourceFiles"][$norm]["PayloadSha256"] = $computedPayloadHash
            }
        }
    }

    return $computedPayloadHash
}

function Get-DirectoryStateHash {
    <#
    .SYNOPSIS
        Computes a deterministic SHA-256 state fingerprint of a directory based on
        its immediate files' names, byte lengths, and LastWriteTimeUtc ticks.
    #>
    param(
        [Parameter()][AllowNull()][object[]]$Files
    )

    if ($null -eq $Files -or $Files.Count -eq 0) {
        return (Get-StringSha256 -InputText "EMPTY_DIR")
    }

    $sortedEntries = @(
        $Files |
        Sort-Object -Property @{ Expression = { $_.Name.ToLowerInvariant() } } |
        ForEach-Object {
            "{0}|{1}|{2}" -f $_.Name.ToLowerInvariant(), $_.Length, $_.LastWriteTimeUtc.Ticks
        }
    )

    $manifest = $sortedEntries -join "`n"
    return (Get-StringSha256 -InputText $manifest)
}

function New-EmptySorterDatabase {
    <#
    .SYNOPSIS
        Creates an empty in-memory database structure with case-insensitive dictionaries.
    #>
    return [ordered]@{
        Version                = 1
        UpdatedUtc             = [datetime]::UtcNow.ToString("o")
        SourceDirectories      = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        SourceFiles            = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        DestinationDirectories = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        DestinationFiles       = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
}

function Import-SorterDatabase {
    <#
    .SYNOPSIS
        Loads the persistent JSON database from disk into case-insensitive hashtables.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$DbFilePath,
        [Parameter(Mandatory = $true)][string]$LogFile
    )

    $db = New-EmptySorterDatabase
    if (-not (Test-Path -LiteralPath $DbFilePath -PathType Leaf)) {
        return $db
    }

    try {
        $rawJson = Get-Content -LiteralPath $DbFilePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($rawJson)) {
            return $db
        }

        $parsed = ConvertFrom-Json -InputObject $rawJson
        if ($null -eq $parsed) {
            return $db
        }

        foreach ($sectionName in @("SourceDirectories", "SourceFiles", "DestinationDirectories", "DestinationFiles")) {
            $prop = $parsed.PSObject.Properties[$sectionName]
            if ($null -ne $prop -and $null -ne $prop.Value) {
                $targetTable = $db[$sectionName]
                foreach ($entryProp in $prop.Value.PSObject.Properties) {
                    $normKey = Get-NormalizedPath -Path $entryProp.Name
                    $valObj = $entryProp.Value
                    $entryHash = @{}
                    foreach ($subProp in $valObj.PSObject.Properties) {
                        $entryHash[$subProp.Name] = $subProp.Value
                    }
                    $targetTable[$normKey] = $entryHash
                }
            }
        }
    }
    catch {
        Write-SorterLog -Message "Failed to load database '$DbFilePath' ($($_.Exception.Message)). Rebuilding database in memory." -Level "WARN" -LogFile $LogFile
        return (New-EmptySorterDatabase)
    }

    return $db
}

function Export-SorterDatabase {
    <#
    .SYNOPSIS
        Persists the in-memory database structure to a UTF-8 JSON file on disk.
    #>
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Database,
        [Parameter(Mandatory = $true)][string]$DbFilePath,
        [Parameter(Mandatory = $true)][string]$LogFile
    )

    try {
        $Database["UpdatedUtc"] = [datetime]::UtcNow.ToString("o")
        $parentDir = Split-Path -Path $DbFilePath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir -PathType Container)) {
            New-Item -Path $parentDir -ItemType Directory -Force -WhatIf:$false | Out-Null
        }

        $jsonOutput = ConvertTo-Json -InputObject $Database -Depth 6
        [System.IO.File]::WriteAllText($DbFilePath, $jsonOutput, [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-SorterLog -Message "Failed to save database to '$DbFilePath': $($_.Exception.Message)" -Level "WARN" -LogFile $LogFile
    }
}

function Get-CachedFileHash {
    <#
    .SYNOPSIS
        Computes and caches the SHA-256 hash of a file to avoid redundant disk reads,
        updating both the session cache and the persistent database if present.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][hashtable]$Cache,
        [Parameter()][System.Collections.IDictionary]$Database = $null,
        [Parameter()][System.Collections.IDictionary]$Stats = $null,
        [Parameter()][System.IO.FileInfo]$FileItem = $null,
        [Parameter()][string]$ReferenceRoot = ""
    )

    $norm = Get-NormalizedPath -Path $FilePath
    if ($Cache.ContainsKey($norm) -and -not [string]::IsNullOrWhiteSpace($Cache[$norm])) {
        return $Cache[$norm]
    }

    $preKey = "__precomputed__:$norm"
    if ($Cache.ContainsKey($preKey) -and -not [string]::IsNullOrWhiteSpace($Cache[$preKey])) {
        $computedFileSha = [string]$Cache[$preKey]
        $Cache[$norm] = $computedFileSha
        if ($null -ne $Database -and $null -ne $FileItem) {
            [void](Set-DatabaseFileRecord -Database $Database -FileItem $FileItem -ReferenceRoot $ReferenceRoot -Sha256 $computedFileSha)
        }
        return $computedFileSha
    }

    if ($null -ne $Database) {
        foreach ($tableName in @("DestinationFiles", "SourceFiles")) {
            if ($Database[$tableName].ContainsKey($norm)) {
                $rec = $Database[$tableName][$norm]
                if ($rec.ContainsKey("Sha256") -and -not [string]::IsNullOrWhiteSpace($rec["Sha256"])) {
                    $isMatch = $true
                    if ($null -ne $FileItem -and $rec.ContainsKey("Length") -and $rec.ContainsKey("LastWriteTimeUtcTicks")) {
                        if ($rec["Length"] -ne $FileItem.Length -or $rec["LastWriteTimeUtcTicks"] -ne $FileItem.LastWriteTimeUtc.Ticks) {
                            $isMatch = $false
                        }
                    }
                    if ($isMatch) {
                        $Cache[$norm] = [string]$rec["Sha256"]
                        if ($null -ne $Stats) { $Stats.CacheHits++ }
                        return $Cache[$norm]
                    }
                }
            }
        }
    }

    $computedFileSha = (Get-FileHash -LiteralPath $norm -Algorithm SHA256).Hash
    $Cache[$norm] = $computedFileSha

    if ($null -ne $Database) {
        if ($null -ne $FileItem) {
            [void](Set-DatabaseFileRecord -Database $Database -FileItem $FileItem -ReferenceRoot $ReferenceRoot -Sha256 $computedFileSha)
        }
        else {
            if ($Database["DestinationFiles"].ContainsKey($norm)) {
                $Database["DestinationFiles"][$norm]["Sha256"] = $computedFileSha
            }
            if ($Database["SourceFiles"].ContainsKey($norm)) {
                $Database["SourceFiles"][$norm]["Sha256"] = $computedFileSha
            }
        }
    }

    return $Cache[$norm]
}

function Sync-DestinationMonthDirectory {
    <#
    .SYNOPSIS
        Synchronizes a destination month directory (\YYYY\YYYY-MM) with the in-memory
        and persistent database using directory state hashing.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$MonthDirPath,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Database,
        [Parameter(Mandatory = $true)][hashtable]$SessionMonthIndex,
        [Parameter(Mandatory = $true)][hashtable]$HashCache,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Stats
    )

    $normMonthDir = Get-NormalizedPath -Path $MonthDirPath
    if ($SessionMonthIndex.ContainsKey($normMonthDir)) {
        return ,$SessionMonthIndex[$normMonthDir]
    }

    $monthEntries = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $normMonthDir -PathType Container)) {
        $SessionMonthIndex[$normMonthDir] = $monthEntries
        return ,$monthEntries
    }

    $diskFiles = @(Get-ChildItem -LiteralPath $normMonthDir -File -ErrorAction SilentlyContinue)
    $currentDirHash = Get-DirectoryStateHash -Files $diskFiles

    $dbDestDirs = $Database["DestinationDirectories"]
    $dbDestFiles = $Database["DestinationFiles"]

    $isDirHashMatch = $false
    if ($dbDestDirs.ContainsKey($normMonthDir)) {
        $savedDirInfo = $dbDestDirs[$normMonthDir]
        if ($savedDirInfo.ContainsKey("DirectoryHash") -and $savedDirInfo["DirectoryHash"] -eq $currentDirHash) {
            $isDirHashMatch = $true
            $Stats.DirectoryHashHits++
        }
    }

    $currentFileSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($df in $diskFiles) {
        $normDf = Get-NormalizedPath -Path $df.FullName
        [void]$currentFileSet.Add($normDf)

        $cachedSha = $null
        $cachedPayloadSha = $null
        $cachedPerceptualHash = $null
        if ($dbDestFiles.ContainsKey($normDf)) {
            $prevRec = $dbDestFiles[$normDf]
            if ($prevRec["Length"] -eq $df.Length -and $prevRec["LastWriteTimeUtcTicks"] -eq $df.LastWriteTimeUtc.Ticks) {
                if ($prevRec.ContainsKey("Sha256") -and -not [string]::IsNullOrWhiteSpace($prevRec["Sha256"])) {
                    $cachedSha = [string]$prevRec["Sha256"]
                }
                if ($prevRec.ContainsKey("PayloadSha256") -and -not [string]::IsNullOrWhiteSpace($prevRec["PayloadSha256"])) {
                    $cachedPayloadSha = [string]$prevRec["PayloadSha256"]
                }
                if ($prevRec.ContainsKey("PerceptualHash") -and -not [string]::IsNullOrWhiteSpace($prevRec["PerceptualHash"])) {
                    $cachedPerceptualHash = [string]$prevRec["PerceptualHash"]
                }
            }
        }

        $dbDestFiles[$normDf] = @{
            Name                  = $df.Name
            DirectoryPath         = $normMonthDir
            Length                = [long]$df.Length
            LastWriteTimeUtcTicks = [long]$df.LastWriteTimeUtc.Ticks
            Sha256                = $cachedSha
            PayloadSha256         = $cachedPayloadSha
            PerceptualHash        = $cachedPerceptualHash
        }

        $monthEntries.Add([PSCustomObject]@{
            FullName = $normDf
            Name     = $df.Name
            Length   = [long]$df.Length
        })
    }

    # Prune deleted files that previously belonged to this destination month directory
    if (-not $isDirHashMatch) {
        $staleKeys = @(
            $dbDestFiles.Keys | Where-Object {
                $rec = $dbDestFiles[$_]
                $rec.ContainsKey("DirectoryPath") -and
                ($rec["DirectoryPath"] -eq $normMonthDir) -and
                (-not $currentFileSet.Contains($_))
            }
        )
        foreach ($sk in $staleKeys) {
            $dbDestFiles.Remove($sk)
        }
    }

    $dbDestDirs[$normMonthDir] = @{
        DirectoryHash = $currentDirHash
        FileCount     = $diskFiles.Count
        UpdatedUtc    = [datetime]::UtcNow.ToString("o")
    }

    $SessionMonthIndex[$normMonthDir] = $monthEntries
    return ,$monthEntries
}

function Update-DirectoryHashesOnDisk {
    <#
    .SYNOPSIS
        Recomputes and updates the DirectoryHash for all modified source and destination
        directories after move/copy operations complete.
    #>
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Database,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$SourceDirsToRefresh,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]]$DestDirsToRefresh,
        [Parameter()][string[]]$NormalizedExtensions = @()
    )

    foreach ($srcDir in $SourceDirsToRefresh) {
        if (Test-Path -LiteralPath $srcDir -PathType Container) {
            $remainingFiles = @(
                Get-ChildItem -LiteralPath $srcDir -File -ErrorAction SilentlyContinue |
                Where-Object {
                    $NormalizedExtensions.Count -eq 0 -or
                    ($NormalizedExtensions -contains $_.Extension.ToLowerInvariant())
                }
            )
            $newHash = Get-DirectoryStateHash -Files $remainingFiles
            $Database["SourceDirectories"][$srcDir] = @{
                DirectoryHash = $newHash
                FileCount     = $remainingFiles.Count
                UpdatedUtc    = [datetime]::UtcNow.ToString("o")
            }
        }
        else {
            $Database["SourceDirectories"].Remove($srcDir)
        }
    }

    foreach ($dstDir in $DestDirsToRefresh) {
        if (Test-Path -LiteralPath $dstDir -PathType Container) {
            $destFiles = @(Get-ChildItem -LiteralPath $dstDir -File -ErrorAction SilentlyContinue)
            $newHash = Get-DirectoryStateHash -Files $destFiles
            $Database["DestinationDirectories"][$dstDir] = @{
                DirectoryHash = $newHash
                FileCount     = $destFiles.Count
                UpdatedUtc    = [datetime]::UtcNow.ToString("o")
            }
        }
        else {
            $Database["DestinationDirectories"].Remove($dstDir)
        }
    }
}
