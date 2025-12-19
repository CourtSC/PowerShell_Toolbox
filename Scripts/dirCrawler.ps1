function Get-DirTreeSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [switch]$Recurse,

        [string]$ExportPath = "$env:OneDrive\Documents\DirectoryTreeSize.xlsx"
    )

    begin {

        if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
            Write-Host 'ImportExcel module not found. Installing...'
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
            Write-Host 'ImportExcel module installed.'
        } else { 
            if ( -not (Get-Module -Name ImportExcel)) {
                Import-Module ImportExcel 
            }
        }

        $ErrorActionPreference = 'Stop'

        # Start runtime stopwatch
        $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        # Normalize $Path once – no trailing backslash
        $Path = (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')


        $output = New-Object System.Collections.Generic.List[object]

        function Stop-ExcelProcess {
            <#
                .SYNOPSIS
                Stops any running Microsoft Excel processes, with optional confirmation and WhatIf support.

                .DESCRIPTION
                Stop-ExcelProcess detects running instances of Microsoft Excel and stops them. By default, if Excel is running,
                the function prompts for confirmation using ShouldContinue. You can bypass the prompt with -Force, or rely on
                PowerShell’s standard -WhatIf / -Confirm behavior via SupportsShouldProcess.

                The function returns a Boolean indicating the outcome:
                - $true  : Excel was not running or was successfully stopped.
                - $false : The user cancelled or an error occurred while stopping Excel.

                .PARAMETER Force
                Skips the interactive confirmation prompt and attempts to stop Excel immediately. Still respects -WhatIf/-Confirm.

                .INPUTS
                None. You cannot pipe input to this function.

                .OUTPUTS
                System.Boolean
                Returns $true on success (or when Excel is already closed), $false on cancellation or error.

                .EXAMPLE
                PS> Stop-ExcelProcess
                Prompts to close running Excel processes, then stops them if confirmed.

                .EXAMPLE
                PS> Stop-ExcelProcess -Force
                Immediately stops any running Excel processes without prompting.

                .EXAMPLE
                PS> Stop-ExcelProcess -WhatIf
                Shows what would happen if the function ran, without making changes.

                .NOTES
                Requires Windows when targeting Microsoft Excel as a desktop application.
                Uses Get-Process 'EXCEL' and Stop-Process -Force.
                Integrates with -WhatIf and -Confirm via SupportsShouldProcess.

                .LINK
                about_Comment_Based_Help
                .LINK
                about_CommonParameters
            #>

            [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
            [OutputType([bool])]
            param(
                [Parameter()]
                [switch]$Force
            )

            # Detect Excel processes safely
            $procs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue

            if (-not $procs) {
                Write-Verbose 'No Excel processes found.'
                return $true
            }

            Write-Verbose ('Detected {0} Excel process(es): {1}' -f $procs.Count, ($procs.Id -join ', '))

            # Optional friendly prompt unless -Force
            if (-not $Force) {
                $caption = 'Excel must be closed before continuing.'
                $message = "Found $($procs.Count) Excel process(es). Close them now?"
                if (-not $PSCmdlet.ShouldContinue($message, $caption)) {
                    Write-Verbose 'Operation cancelled by user.'
                    return $false
                }
            }

            # Respect -WhatIf / -Confirm
            if ($PSCmdlet.ShouldProcess("Excel ($($procs.Count))", 'Stop-Process -Force')) {
                try {
                    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
                    Write-Verbose 'Successfully closed Excel.'
                    return $true
                } catch {
                    Write-Error -ErrorRecord $_
                    return $false
                }
            }
        }
    }

    process {
        if (-not $Recurse) {
            # Single directory stats only
            $files = Get-ChildItem -LiteralPath $Path -File -ErrorAction Stop

            $fileStats = $files | Measure-Object -Property Length -Sum
            $fileCount = $fileStats.Count

            $directoryCount = (Get-ChildItem -LiteralPath $Path -Directory -ErrorAction Stop | Measure-Object).Count

            $sizeMB = if ($fileStats.Sum) { [math]::Round($fileStats.Sum / 1MB, 3) } else { 0 }

            $directoryOwner = (Get-Acl -LiteralPath $Path).Owner
            $depth = 0

            if ($files.Count -gt 0) {
                $lastWrite = ($files | Measure-Object LastWriteTime -Maximum).Maximum
                $lastAccess = ($files | Measure-Object LastAccessTime -Maximum).Maximum

                $lastModifiedDate = $lastWrite.ToString('MM/dd/yyyy HH:mm')
                $lastAccessDate = $lastAccess.ToString('MM/dd/yyyy HH:mm')
            }

            $output.Add([pscustomobject]@{
                    Path           = $Path
                    Parent         = $null
                    Depth          = $depth
                    Owner          = $directoryOwner
                    FileCount      = $fileCount
                    DirectoryCount = $directoryCount
                    DirSizeInMB    = $sizeMB
                    LastModified   = $lastModifiedDate
                    LastAccessed   = $lastAccessDate
                }) | Out-Null

            return  # end block will handle runtime + return value
        }

        # ---- Recursive version below ----

        # Grab ALL directories once
        $dirs = Get-ChildItem -LiteralPath $Path -Directory -Recurse -ErrorAction Stop

        # Normalize directory paths (no trailing slashes)
        $rootPath = $Path.TrimEnd('\')
        $allDirs = @($rootPath) + ($dirs | ForEach-Object { $_.FullName.TrimEnd('\') })

        # Grab ALL files once
        $allFiles = Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction Stop

        # Group files by normalized directory name
        $filesByDir = $allFiles | ForEach-Object {
            $_ | Add-Member -NotePropertyName NormalizedDirectory `
                -NotePropertyValue ($_.DirectoryName.TrimEnd('\')) -PassThru
        } | Group-Object -Property NormalizedDirectory -AsHashTable -AsString


        foreach ($dirPath in $allDirs) {

            $dirFiles = if ($filesByDir.ContainsKey($dirPath)) { $filesByDir[$dirPath] } else { @() }

            $fileStats = $dirFiles | Measure-Object -Property Length -Sum
            $fileCount = $fileStats.Count
            $sizeMB = if ($fileStats.Sum) { [math]::Round($fileStats.Sum / 1MB, 3) } else { 0 }

            $childDirCount = ($dirs | Where-Object {
                    ( $_.FullName.TrimEnd('\') | Split-Path -Parent ) -eq $dirPath
                } | Measure-Object).Count

            $owner = (Get-Acl -LiteralPath $dirPath).Owner

            if ($dirPath -eq $rootPath) {
                $depth = 0
                $parent = $null
            } else {
                $relative = $dirPath.Substring($rootPath.Length).TrimStart('\')
                $depth = ($relative -split '\\').Count

                $parent = Split-Path -Path $dirPath -Parent
                $parent = $parent.TrimEnd('\')
                if ($parent -eq '') { $parent = $null }
            }

            # Avoid leaking values from previous iterations
            $lastModifiedDate = $null
            $lastAccessDate = $null

            if ($dirFiles.Count -gt 0) {
                $lastWrite = ($dirFiles | Measure-Object LastWriteTime -Maximum).Maximum
                $lastAccess = ($dirFiles | Measure-Object LastAccessTime -Maximum).Maximum

                $lastModifiedDate = $lastWrite.ToString('MM/dd/yyyy HH:mm')
                $lastAccessDate = $lastAccess.ToString('MM/dd/yyyy HH:mm')
            }

            $output.Add([pscustomobject]@{
                    Path           = $dirPath
                    Parent         = $parent
                    Depth          = $depth
                    Owner          = $owner
                    FileCount      = $fileCount
                    DirectoryCount = $childDirCount
                    DirSizeInMB    = $sizeMB
                    LastModified   = $lastModifiedDate
                    LastAccessed   = $lastAccessDate
                }) | Out-Null
        }



        # Compute TotalSizeInMB bottom-up without O(n²)
        $byPath = @{}
        $children = @{}

        foreach ($item in $output) {
            $byPath[$item.Path] = $item
            if ($item.Parent) {
                if (-not $children.ContainsKey($item.Parent)) {
                    $children[$item.Parent] = New-Object System.Collections.Generic.List[object]
                }
                $children[$item.Parent].Add($item) | Out-Null
            }
        }

        foreach ($item in $output | Sort-Object Depth -Descending) {
            $total = [decimal]$item.DirSizeInMB
            if ($children.ContainsKey($item.Path)) {
                foreach ($child in $children[$item.Path]) {
                    $total += [decimal]$child.TotalSizeInMB
                }
            }
            $item | Add-Member -NotePropertyName TotalSizeInMB -NotePropertyValue $total -Force
        }

    }
    
    end {
        if ($script:Stopwatch) {
            $script:Stopwatch.Stop()
            $elapsed = $script:Stopwatch.Elapsed
            
            # Verbose message
            Write-Verbose ('Get-DirTreeSize completed in {0:hh\:mm\:ss\.fff}' -f $elapsed)
            
            # Add runtime to root row (optional)
            $root = $output | Where-Object { $_.Path -eq $Path } | Select-Object -First 1
            if ($null -ne $root) {
                $root | Add-Member -NotePropertyName TotalRuntimeSeconds -NotePropertyValue ([math]::Round($elapsed.TotalSeconds, 3)) -Force
            }
        }

        # Export once at the end
        if ($output.Count -gt 0 -and $Recurse) {
            Stop-ExcelProcess -Force | Out-Null
            Start-Sleep -Seconds 3
            $leaf = Split-Path -Path $Path -Leaf
            $output | Export-Excel -Path $ExportPath -WorksheetName $leaf -TableName "$($leaf)_Data" -AutoSize -ClearSheet
        }

        # Non-recursive calls return objects; recursive already exported, but you might still want the objects:
        if ($output -and -not $Recurse) {
            return $output
        }
    }
}