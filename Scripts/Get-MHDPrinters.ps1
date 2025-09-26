function Get-MHDPrinters {
    <#
.SYNOPSIS
Gets printer information from a set of MHD print servers.

.DESCRIPTION
Queries one or more print servers for one or more specified printer names.
By default, queries the MHD print servers (ADCVPRNMHDMS001–006, 020–024, 030–031, 040–041, 050–052, 060–062, 070–071, 080–081).

Runs work in parallel using thread jobs (fast and progress-safe). Progress is reported from the parent runspace if `-Verbose` is set.
Returns structured objects with server, printer, driver, port, and AD group info.

Port data is served from a session-wide cache (`script:PortCache`) to avoid repeated remote calls.
The cache is built on-demand (first run) or when `-RebuildPortCache` is specified. Jobs **never** call `Get-PrinterPort` directly.
If a specific port is missing from the cache, those port fields will be `$null` until you rebuild the cache.

.PARAMETER Printers
One or more printer names to look up. Accepts pipeline input.

.PARAMETER Servers
One or more print server names to query. Defaults to the MHD fleet listed above.

.PARAMETER ThrottleLimit
Maximum number of parallel jobs. Default: 12.

.PARAMETER NoProgress
Suppress the Write-Progress UI. Note: progress is shown only when -Verbose is used; -NoProgress suppresses it even then.

.PARAMETER NoPort
Skips using or building the port cache and omits port info (HostAddress, PortDescription, SNMPEnabled = $null).

.PARAMETER RebuildPortCache
Rebuilds the session-wide port cache before querying (forces fresh `Get-PrinterPort` calls in the parent runspace).

.PARAMETER ServerTimeoutSec
Timeout (in seconds) for each server/printer job. Default: 60.

.EXAMPLE
PS> Get-MHDPrinters -Printers 'HP123','CanonX' -RebuildPortCache

.EXAMPLE
PS> Get-MHDPrinters -NoPort -ThrottleLimit 24

.EXAMPLE
PS> 'HP123' | Get-MHDPrinters -Verbose

.NOTES
- Port cache variable: script:PortCache  (Hashtable: Server -> (Hashtable: PortName -> PrinterPort object))
- Cache lifetime: current PowerShell session.
- Requires the PrintManagement and ActiveDirectory modules.
.LINK
Get-Printer
#>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$Printers,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]$Servers = @(
            'ADCVPRNMHDMS001', 'ADCVPRNMHDMS002', 'ADCVPRNMHDMS003', 'ADCVPRNMHDMS004', 'ADCVPRNMHDMS005', 'ADCVPRNMHDMS006',
            'ADCVPRNMHDMS020', 'ADCVPRNMHDMS021', 'ADCVPRNMHDMS022', 'ADCVPRNMHDMS023', 'ADCVPRNMHDMS024',
            'ADCVPRNMHDMS030', 'ADCVPRNMHDMS031',
            'ADCVPRNMHDMS040', 'ADCVPRNMHDMS041',
            'ADCVPRNMHDMS050', 'ADCVPRNMHDMS051', 'ADCVPRNMHDMS052',
            'ADCVPRNMHDMS060', 'ADCVPRNMHDMS061', 'ADCVPRNMHDMS062',
            'ADCVPRNMHDMS070', 'ADCVPRNMHDMS071',
            'ADCVPRNMHDMS080', 'ADCVPRNMHDMS081'
        ),

        [ValidateRange(1, 256)]
        [int]$ThrottleLimit = 12,

        [switch]$NoProgress,

        [switch]$NoPort,

        [switch]$RebuildPortCache,

        [ValidateRange(5, 600)]
        [int]$ServerTimeoutSec = 60
    )

    begin {
        # Show progress only when -Verbose is supplied; allow -NoProgress to suppress even then
        $ShowProgress = ($PSBoundParameters['Verbose']) -and (-not $NoProgress)

        Import-Module PrintManagement -ErrorAction Stop
        Import-Module ActiveDirectory -ErrorAction Stop
        
        if ($PSBoundParameters['Verbose'] -or $RebuildPortCache -or (-not $script:PortCache)) {
            Write-Host 'Building printer port cache... please be patient.' -ForegroundColor Yellow
        }
        Write-Host 'Tip: Use -RebuildPortCache if you suspect ports have changed on a server.' -ForegroundColor DarkGray

        # Ensure global (script:) port cache exists; (re)build if requested and we're not skipping ports
        if (-not $NoPort) {
            if (-not (Get-Variable -Name PortCache -Scope Script -ErrorAction SilentlyContinue)) {
                Set-Variable -Name PortCache -Scope Script -Value (@{}) -Option None
            }
            if ($RebuildPortCache) {
                $script:PortCache = @{}
            }

            foreach ($srv in $Servers) {
                if (-not $script:PortCache.ContainsKey($srv)) {
                    # Build per-server cache ONCE, here in parent runspace
                    try {
                        $serverPorts = @{}
                        foreach ($prt in (Get-PrinterPort -ComputerName $srv -ErrorAction Stop)) {
                            $serverPorts[$prt.Name] = $prt
                        }
                        $script:PortCache[$srv] = $serverPorts
                        Write-Verbose "Cached $($serverPorts.Count) ports for $srv."
                    } catch {
                        # If we can't cache this server now, store an empty map to avoid trying in jobs
                        $script:PortCache[$srv] = @{}
                        Write-Verbose "Failed to cache ports for $($srv): $($_.Exception.Message)"
                    }
                }
            }
        }

        # Create a snapshot for jobs (jobs can't see parent scopes; we pass what they need)
        $PortCacheSnapshot = if ($NoPort) { @{} } else { $script:PortCache }

        $jobScriptEnumerateServer = {
            param($Server, $TimeoutSec, $SkipPort, $PortMapForServer)

            # Import-Module PrintManagement -ErrorAction Stop
            # Import-Module ActiveDirectory -ErrorAction Stop

            try {
                $printers = Get-Printer -ComputerName $Server -ErrorAction Stop
                foreach ($p in $printers) {
                    # Port lookup strictly from cache snapshot; jobs NEVER call Get-PrinterPort
                    $port = $null
                    if (-not $SkipPort -and $PortMapForServer) {
                        if ($PortMapForServer.ContainsKey($p.PortName)) {
                            $port = $PortMapForServer[$p.PortName]
                        }
                    }

                    $filterString = "*-PRN-$($p.Name)*"
                    $stdGroup = $null
                    $defGroup = $null
                    try {
                        $groups = Get-ADGroup -Filter "Name -like '$filterString'" -ErrorAction Stop
                        foreach ($g in $groups) {
                            if ($g.Name -like '*-PRN-*' -and $g.Name -like '*-DEF') { $defGroup = $g.Name }
                            elseif ($g.Name -like '*-PRN-*') { $stdGroup = $g.Name }
                        }
                    } catch { }

                    [pscustomobject]@{
                        Server          = $p.ComputerName
                        Name            = $p.Name
                        Status          = $p.PrinterStatus
                        StandardGroup   = $stdGroup
                        DefaultGroup    = $defGroup
                        DriverName      = $p.DriverName
                        PortName        = $p.PortName
                        HostAddress     = $port.PrinterHostAddress
                        PortDescription = $port.Description
                        SNMPEnabled     = $port.SNMPEnabled
                        Shared          = $p.Shared
                        Published       = $p.Published
                        Comment         = $p.Comment
                        Location        = $p.Location
                    }
                }
            } catch {
                Write-Verbose "[$Server] Failed to enumerate printers: $($_.Exception.Message)"
            }
        }

        $jobScriptPrinterOnServer = {
            param($Server, $Printer, $TimeoutSec, $SkipPort, $PortMapForServer)

            # Import-Module PrintManagement -ErrorAction Stop
            # Import-Module ActiveDirectory -ErrorAction Stop

            try {
                $p = Get-Printer -ComputerName $Server -Name $Printer -ErrorAction Stop
            } catch {
                Write-Verbose "Printer '$Printer' not found on $Server. ($($_.Exception.Message))"
                return
            }

            # Port lookup strictly from cache snapshot; jobs NEVER call Get-PrinterPort
            $port = $null
            if (-not $SkipPort -and $PortMapForServer) {
                if ($PortMapForServer.ContainsKey($p.PortName)) {
                    $port = $PortMapForServer[$p.PortName]
                }
            }

            $filterString = "*-PRN-$($p.Name)*"
            $stdGroup = $null
            $defGroup = $null
            try {
                $groups = Get-ADGroup -Filter "Name -like '$filterString'" -ErrorAction Stop
                foreach ($g in $groups) {
                    if ($g.Name -like '*-PRN-*' -and $g.Name -like '*-DEF') { $defGroup = $g.Name }
                    elseif ($g.Name -like '*-PRN-*') { $stdGroup = $g.Name }
                }
            } catch { }

            [pscustomobject]@{
                Server          = $p.ComputerName
                Name            = $p.Name
                Status          = $p.PrinterStatus
                StandardGroup   = $stdGroup
                DefaultGroup    = $defGroup
                DriverName      = $p.DriverName
                PortName        = $p.PortName
                HostAddress     = $port.PrinterHostAddress
                PortDescription = $port.Description
                SNMPEnabled     = $port.SNMPEnabled
                Shared          = $p.Shared
                Published       = $p.Published
                Comment         = $p.Comment
                Location        = $p.Location
            }
        }

        $work = New-Object System.Collections.Generic.List[object]
    }

    process {
        if ($Printers) {
            foreach ($printer in $Printers) {
                foreach ($srv in $Servers) {
                    $work.Add([pscustomobject]@{ Type = 'Specific'; Server = $srv; Printer = $printer })
                }
            }
        } else {
            foreach ($srv in $Servers) {
                $work.Add([pscustomobject]@{ Type = 'Enumerate'; Server = $srv })
            }
        }
    }

    end {
        $jobs = @()

        $total = $work.Count
        $started = 0
        $completed = 0

        if ($ShowProgress) {
            Write-Progress -Id 1 -Activity 'Querying print servers' -Status 'Starting...' -PercentComplete 0
        }

        while ($started -lt $total -or ($jobs | Where-Object State -In 'Running', 'NotStarted')) {
            while ($started -lt $total -and ($jobs | Where-Object State -In 'Running', 'NotStarted').Count -lt $ThrottleLimit) {
                $item = $work[$started]
                $serverPortMap = if ($NoPort) { $null } else { $PortCacheSnapshot[$($item.Server)] }

                if ($item.Type -eq 'Enumerate') {
                    $jobs += Start-ThreadJob `
                        -ScriptBlock $jobScriptEnumerateServer `
                        -ArgumentList $item.Server, $ServerTimeoutSec, $NoPort.IsPresent, $serverPortMap `
                        -Name "Enum:$($item.Server)" `
                        -StreamingHost $Host
                } else {
                    $jobs += Start-ThreadJob `
                        -ScriptBlock $jobScriptPrinterOnServer `
                        -ArgumentList $item.Server, $item.Printer, $ServerTimeoutSec, $NoPort.IsPresent, $serverPortMap `
                        -Name "Get:$($item.Server):$($item.Printer)" `
                        -StreamingHost $Host
                }
                $started++
            }

            $ready = $jobs | Where-Object State -EQ 'Completed'
            if ($ready) {
                $ready | Receive-Job -Keep | Write-Output
                $completed += $ready.Count
                if ($ShowProgress) {
                    $pct = [math]::Min(100, [math]::Round(($completed / [math]::Max(1, $total)) * 100, 2))
                    $status = "$completed of $total finished"
                    Write-Progress -Id 1 -Activity 'Querying print servers' -Status $status -PercentComplete $pct -CurrentOperation ($ready[-1].Name)
                }
                $jobs = $jobs | Where-Object State -NE 'Completed'
            }

            Start-Sleep -Milliseconds 150
        }

        if ($jobs) {
            $more = $jobs | Receive-Job -Wait -AutoRemoveJob
            if ($more) { $more | Write-Output }
        }

        if ($ShowProgress) {
            Write-Progress -Id 1 -Activity 'Querying print servers' -Completed
        }
    }
}