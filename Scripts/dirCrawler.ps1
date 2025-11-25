function Get-DirTreeSize {

    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [Parameter()]
        [switch]$Recurse,

        [Parameter()]
        [string]$ExportPath = "$env:OneDrive\Documents\DirectoryTreeSize.xlsx"
    )

    begin {
        $output = @()

        #Enabling long paths in Windows to avoid character limits.
        $longPathsEnabled = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled'
        if (-not $longPathsEnabled) {
            New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Value 1 -PropertyType DWORD -Force | Out-Null 
        }

        #Adding a trailing slash at the end of $path to make it consistent.
        if (-not $Path.EndsWith('\')) {
            $Path = "$Path\"
        }

        #Install or import the ImportExcel module.
        if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
            Write-Host 'ImportExcel module not found. Installing...'
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
            Write-Host 'ImportExcel module installed.'
        } else { Import-Module ImportExcel }
    
        function Stop-ExcelProcess {
    
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
                    $procs | Stop-Process -Force -ErrorAction Stop
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

        if (-not $PSBoundParameters.ContainsKey('Recurse')) {
            $files = Get-ChildItem -Path $Path -File -ErrorAction Stop
            $fileStats = $files | Measure-Object -Property Length -Sum
            $fileCount = $fileStats.Count
            $directoryCount = Get-ChildItem -Path $Path -Directory | Measure-Object | Select-Object -ExpandProperty Count
            $sizeMB = '{0:F3}' -f ($FileStats.Sum / 1MB) -as [decimal]
            $directoryOwner = (Get-Acl -Path $Path).Owner
            $depth = 0

            if ($files) {
                $lastModifiedDate = Get-Date ($files | Sort-Object LastWriteTime | Select-Object -ExpandProperty LastWriteTime -Last 1) -Format 'MM/dd/yyyy HH:mm'
                $lastAccessDate = Get-Date ($files | Sort-Object LastAccessTime | Select-Object -ExpandProperty LastAccessTime -Last 1) -Format 'MM/dd/yyyy HH:mm'

            }

            $output = [PSCustomObject]@{
                Path           = $Path
                Parent         = $null
                Depth          = $depth
                Owner          = $directoryOwner
                FileCount      = $FileCount
                DirectoryCount = $DirectoryCount
                DirSizeInMB    = $SizeMB
                LastModified   = $lastModifiedDate
                LastAccessed   = $lastAccessDate

            }
        }

        if ($PSBoundParameters.ContainsKey('Recurse')) {
            #Emit stats for the top-level directory non-recursively.
            $output += Get-DirTreeSize -Path $Path

            #Construct a list of directories to accumulate directory stats.
            $directoryList = Get-ChildItem -Path $Path -Directory -Recurse | Select-Object -ExpandProperty FullName

            #Recursive logic.
            if ($directoryList) {
                $output += foreach ($dir in $directoryList) {
                    $files = Get-ChildItem -Path $dir -File -ErrorAction Stop
                    $fileStats = $files | Measure-Object -Property Length -Sum
                    $fileCount = $fileStats.Count
                    $directoryCount = Get-ChildItem -Path $dir -Directory | Measure-Object | Select-Object -ExpandProperty Count
                    $sizeMB = '{0:F3}' -f ($FileStats.Sum / 1MB) -as [decimal]
                    $directoryOwner = (Get-Acl -Path $dir).Owner
                    $depth = ($dir.Replace($Path, '.\') -split '\\').Count - 1
                    $parent = Split-Path $dir

                    if ($files) {
                        $lastModifiedDate = Get-Date ($files | Sort-Object LastWriteTime | Select-Object -ExpandProperty LastWriteTime -Last 1) -Format 'MM/dd/yyyy HH:mm'
                        $lastAccessDate = Get-Date ($files | Sort-Object LastAccessTime | Select-Object -ExpandProperty LastAccessTime -Last 1) -Format 'MM/dd/yyyy HH:mm'
                    }

                    [PSCustomObject]@{
                        Path           = $dir
                        Parent         = $parent
                        Depth          = $depth
                        Owner          = $directoryOwner
                        FileCount      = $FileCount
                        DirectoryCount = $DirectoryCount
                        DirSizeInMB    = $SizeMB
                        LastModified   = $lastModifiedDate
                        LastAccessed   = $lastAccessDate
                    }

                    #clearing variables
                    $files = $null
                    $fileStats = $null
                    $fileCount = $null
                    $directoryCount = $null
                    $sizeMB = $null
                    $directoryOwner = $null
                    $depth = $null
                    $lastModifiedDate = $null
                    $lastAccessDate = $null
                }
            }

            foreach ($obj in $output) {
                $parentPath = $obj.Path
                $children = $output | Where-Object { $_.Path -match $parentPath.Replace('\', '\\') }
                $childSize = ($children | ForEach-Object DirSizeInMB | Measure-Object -Sum).Sum
                $output | Where-Object { $_.Path -eq $parentPath } | ForEach-Object { $_ | Add-Member -NotePropertyName TotalSizeInMB -NotePropertyValue $childSize }
            }
        }
    }

    end {
        if (-not $PSBoundParameters.ContainsKey('Recurse')) {
            return $output
        }
        if ($PSBoundParameters.ContainsKey('Recurse')) {
            Stop-ExcelProcess -Force | Out-Null
            Start-Sleep -Seconds 3
            $leaf = Split-Path -Path $Path -Leaf
            $output | Export-Excel -Path $ExportPath -WorksheetName $leaf -TableName "$($leaf)_Data" -AutoSize -ClearSheet
        }
    }
}