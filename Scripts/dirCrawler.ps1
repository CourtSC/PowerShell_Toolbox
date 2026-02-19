function Get-DirTreeSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [switch]$Recurse,

        [string]$ExportPath = "$env:OneDrive\Documents\DirectoryTreeSize.csv",

        # Optional: don’t scan deeper than this (huge perf win on monster shares)
        [int]$MaxDepth = [int]::MaxValue,

        # Optional: skip owner lookups (ACL calls are slow on shares)
        [switch]$SkipOwner
    )

    $ErrorActionPreference = 'Stop'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $rootPath = (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')

    if (-not $Recurse) {
        $files = Get-ChildItem -LiteralPath $rootPath -File -ErrorAction Stop
        $fileStats = $files | Measure-Object -Property Length -Sum
        $directoryCount = (Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction Stop | Measure-Object).Count

        $lastWrite = if ($files.Count) { ($files | Measure-Object LastWriteTime -Maximum).Maximum } else { $null }
        $lastAccess = if ($files.Count) { ($files | Measure-Object LastAccessTime -Maximum).Maximum } else { $null }

        [pscustomobject]@{
            Path                = $rootPath
            Parent              = $null
            Depth               = 0
            Owner               = if ($SkipOwner) { $null } else { (Get-Acl -LiteralPath $rootPath).Owner }
            FileCount           = $fileStats.Count
            DirectoryCount      = $directoryCount
            DirSizeInMB         = if ($fileStats.Sum) { [math]::Round($fileStats.Sum / 1MB, 3) } else { 0 }
            TotalSizeInMB       = if ($fileStats.Sum) { [math]::Round($fileStats.Sum / 1MB, 3) } else { 0 }
            LastModified        = if ($lastWrite) { $lastWrite.ToString('MM/dd/yyyy HH:mm') } else { $null }
            LastAccessed        = if ($lastAccess) { $lastAccess.ToString('MM/dd/yyyy HH:mm') } else { $null }
            TotalRuntimeSeconds = $null
        }
        return
    }

    # --- Aggregation tables (keyed by normalized dir path) ---
    $sizeBytesByDir = @{}
    $fileCountByDir = @{}
    $lastWriteByDir = @{}
    $lastAccessByDir = @{}
    $dirSet = [System.Collections.Generic.HashSet[string]]::new()


    # Ensure root exists in tables
    $dirSet.Add($rootPath) | Out-Null
    $sizeBytesByDir[$rootPath] = 0L
    $fileCountByDir[$rootPath] = 0

    # --- Stream files once ---
    Get-ChildItem -LiteralPath $rootPath -File -Recurse -Force -ErrorAction Stop | ForEach-Object {
        $dir = $_.DirectoryName.TrimEnd('\')

        # Depth filter (optional)
        if ($MaxDepth -ne [int]::MaxValue) {
            $rel = $dir.Substring($rootPath.Length).TrimStart('\')
            $depth = if ($rel) { ($rel -split '\\').Count } else { 0 }
            if ($depth -gt $MaxDepth) { return }
        }

        $dirSet.Add($dir) | Out-Null

        if (-not $sizeBytesByDir.ContainsKey($dir)) {
            $sizeBytesByDir[$dir] = 0L
            $fileCountByDir[$dir] = 0
        }

        $sizeBytesByDir[$dir] += [int64]$_.Length
        $fileCountByDir[$dir] += 1

        $lw = $_.LastWriteTime
        if (-not $lastWriteByDir.ContainsKey($dir) -or $lw -gt $lastWriteByDir[$dir]) {
            $lastWriteByDir[$dir] = $lw
        }

        $la = $_.LastAccessTime
        if (-not $lastAccessByDir.ContainsKey($dir) -or $la -gt $lastAccessByDir[$dir]) {
            $lastAccessByDir[$dir] = $la
        }
    }

    # --- Enumerate directories (also streaming) ---
    Get-ChildItem -LiteralPath $rootPath -Directory -Recurse -Force -ErrorAction Stop | ForEach-Object {
        $dir = $_.FullName.TrimEnd('\')

        if ($MaxDepth -ne [int]::MaxValue) {
            $rel = $dir.Substring($rootPath.Length).TrimStart('\')
            $depth = if ($rel) { ($rel -split '\\').Count } else { 0 }
            if ($depth -gt $MaxDepth) { return }
        }

        $dirSet.Add($dir) | Out-Null
        if (-not $sizeBytesByDir.ContainsKey($dir)) {
            $sizeBytesByDir[$dir] = 0L
            $fileCountByDir[$dir] = 0
        }
    }

    # Build list of dirs and compute parent/depth
    # Build list of directories in a robust way
    if ($null -eq $dirSet) {
        throw 'Unexpected: dirSet is null'
    }

    # If dirSe# Build list of directories in a robust way
    if ($null -eq $dirSet) {
        throw 'Unexpected: dirSet is null'
    }

    if ($dirSet -is [System.Array]) {
        $allDirs = $dirSet
    } elseif ($dirSet -is [System.Collections.IEnumerable]) {
        try {
            $allDirs = $dirSet.ToArray()
        } catch {
            $allDirs = @($dirSet)
        }

        # Normalize to strings
        $allDirs = $allDirs | ForEach-Object { $_.ToString() }
    } else {
        $allDirs = @($dirSet.ToString())
    }
    $rows = New-Object System.Collections.Generic.List[object]

    # Child mapping for totals and directory counts
    $children = @{}
    foreach ($dir in $allDirs) {
        if ($dir -eq $rootPath) { continue }
        $parent = (Split-Path -Path $dir -Parent).TrimEnd('\')
        if (-not $children.ContainsKey($parent)) {
            $children[$parent] = New-Object System.Collections.Generic.List[string]
        }
        $children[$parent].Add($dir) | Out-Null
    }

    foreach ($dir in $allDirs) {
        $parent = if ($dir -eq $rootPath) { $null } else { (Split-Path -Path $dir -Parent).TrimEnd('\') }

        $rel = if ($dir -eq $rootPath) { '' } else { $dir.Substring($rootPath.Length).TrimStart('\') }
        $depth = if ($rel) { ($rel -split '\\').Count } else { 0 }

        $dirBytes = [int64]$sizeBytesByDir[$dir]
        $dirMB = [math]::Round($dirBytes / 1MB, 3)

        $owner = $null
        if (-not $SkipOwner) {
            try { $owner = (Get-Acl -LiteralPath $dir -ErrorAction Stop).Owner } catch { $owner = $null }
        }

        $rows.Add([pscustomobject]@{
                Path           = $dir
                Parent         = $parent
                Depth          = $depth
                Owner          = $owner
                FileCount      = [int]$fileCountByDir[$dir]
                DirectoryCount = if ($children.ContainsKey($dir)) { $children[$dir].Count } else { 0 }
                DirSizeInMB    = $dirMB
                TotalSizeInMB  = 0.0  # filled later
                LastModified   = if ($lastWriteByDir.ContainsKey($dir)) { $lastWriteByDir[$dir].ToString('MM/dd/yyyy HH:mm') } else { $null }
                LastAccessed   = if ($lastAccessByDir.ContainsKey($dir)) { $lastAccessByDir[$dir].ToString('MM/dd/yyyy HH:mm') } else { $null }
            }) | Out-Null
    }

    # --- Compute TotalSize bottom-up (directories only) ---
    $byPath = @{}
    foreach ($r in $rows) { $byPath[$r.Path] = $r }

    foreach ($r in ($rows | Sort-Object Depth -Descending)) {
        $total = [double]$r.DirSizeInMB
        if ($children.ContainsKey($r.Path)) {
            foreach ($c in $children[$r.Path]) {
                $total += [double]$byPath[$c].TotalSizeInMB
            }
        }
        $r.TotalSizeInMB = [math]::Round($total, 3)
    }

    $sw.Stop()
    $rootRow = $byPath[$rootPath]
    $rootRow | Add-Member -NotePropertyName TotalRuntimeSeconds -NotePropertyValue ([math]::Round($sw.Elapsed.TotalSeconds, 3)) -Force

    # --- CSV export (fast, opens clean in Excel) ---
    $rows |
        Sort-Object Depth, Path |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

    Write-Verbose ('Exported {0} rows to {1}. Runtime: {2}' -f $rows.Count, $ExportPath, $sw.Elapsed)
    $rows
}
 