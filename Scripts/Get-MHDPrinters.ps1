function Get-MHDPrinters {
    <#
.SYNOPSIS
Gets printer information from a set of MHD print servers.

.DESCRIPTION
Queries one or more print servers for one or more specified printer names.
By default, queries the MHD print servers (ADCVPRNMHDMS001–006, 020–024, 030–031, 040–041, 050–052, 060–062, 070–071, 080–081).

Runs work in parallel using thread jobs (fast and progress-safe). Progress is reported from the parent runspace.
Returns structured objects with server, printer, driver, port, and AD group info.

.PARAMETER Printers
One or more printer names to look up. Accepts pipeline input.

.PARAMETER Servers
One or more print server names to query. Defaults to the MHD fleet listed above.

.PARAMETER ThrottleLimit
Maximum number of parallel jobs. Default: 12.

.PARAMETER NoProgress
Suppress the Write-Progress UI. Note: progress is shown only when -Verbose is used; -NoProgress suppresses it even then.

.PARAMETER ServerTimeoutSec
Timeout (in seconds) for each server/printer job. Default: 60.

.EXAMPLE
PS> Get-MHDPrinters -Printers 'HP123','CanonX'

.EXAMPLE
PS> 'HP123' | Get-MHDPrinters -ThrottleLimit 24 -Verbose

.NOTES
Requires the PrintManagement and ActiveDirectory modules.
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

        [ValidateRange(5, 600)]
        [int]$ServerTimeoutSec = 60
    )

    begin {
        # Show progress only when -Verbose is supplied; allow -NoProgress to suppress even then
        $ShowProgress = ($PSBoundParameters['Verbose']) -and (-not $NoProgress)

        # Load modules in parent runspace (also import inside jobs)
        Import-Module PrintManagement -ErrorAction Stop
        Import-Module ActiveDirectory -ErrorAction Stop

        $jobScriptEnumerateServer = {
            param($Server, $TimeoutSec)

            # Ensure modules exist in worker
            # Import-Module PrintManagement -ErrorAction Stop
            # Import-Module ActiveDirectory -ErrorAction Stop

            try {
                # Get all printers once
                $printers = Get-Printer -ComputerName $Server -ErrorAction Stop

                # Build a map of ports once per server
                $ports = @{}
                try {
                    foreach ($prt in (Get-PrinterPort -ComputerName $Server -ErrorAction Stop)) {
                        $ports[$prt.Name] = $prt
                    }
                } catch {
                    # Port enumeration may fail; we'll try per-printer fallback below if needed
                }

                foreach ($p in $printers) {
                    $port = $null
                    if ($ports.ContainsKey($p.PortName)) {
                        $port = $ports[$p.PortName]
                    } else {
                        try {
                            $port = Get-PrinterPort -ComputerName $Server -Name $p.PortName -ErrorAction Stop
                        } catch { }
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
            param($Server, $Printer, $TimeoutSec)

            # Import-Module PrintManagement -ErrorAction Stop
            # Import-Module ActiveDirectory -ErrorAction Stop

            try {
                $p = Get-Printer -ComputerName $Server -Name $Printer -ErrorAction Stop
            } catch {
                Write-Verbose "Printer '$Printer' not found on $Server. ($($_.Exception.Message))"
                return
            }

            $port = $null
            try {
                $port = Get-PrinterPort -ComputerName $Server -Name $p.PortName -ErrorAction Stop
            } catch { }

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

        # Start jobs (bounded by ThrottleLimit using a simple queue)
        $total = $work.Count
        $started = 0
        $completed = 0

        if ($ShowProgress) {
            Write-Progress -Id 1 -Activity 'Querying print servers' -Status 'Starting...' -PercentComplete 0
        }

        while ($started -lt $total -or ($jobs | Where-Object State -In 'Running', 'NotStarted')) {
            # Fill up to throttle
            while ($started -lt $total -and ($jobs | Where-Object State -In 'Running', 'NotStarted').Count -lt $ThrottleLimit) {
                $item = $work[$started]
                if ($item.Type -eq 'Enumerate') {
                    $jobs += Start-ThreadJob -ScriptBlock $jobScriptEnumerateServer -ArgumentList $item.Server, $ServerTimeoutSec -Name "Enum:$($item.Server)" -StreamingHost $Host
                } else {
                    $jobs += Start-ThreadJob -ScriptBlock $jobScriptPrinterOnServer -ArgumentList $item.Server, $item.Printer, $ServerTimeoutSec -Name "Get:$($item.Server):$($item.Printer)" -StreamingHost $Host
                }
                $started++
            }

            # Receive any completed results so far
            $ready = $jobs | Where-Object State -EQ 'Completed'
            if ($ready) {
                $ready | Receive-Job -Keep | Write-Output
                $completed += $ready.Count
                if ($ShowProgress) {
                    $pct = [math]::Min(100, [math]::Round(($completed / [math]::Max(1, $total)) * 100, 2))
                    $status = "$completed of $total finished"
                    Write-Progress -Id 1 -Activity 'Querying print servers' -Status $status -PercentComplete $pct -CurrentOperation ($ready[-1].Name)
                }
                # Clean up completed jobs so we don't count them twice
                $jobs = $jobs | Where-Object State -NE 'Completed'
            }

            Start-Sleep -Milliseconds 150
        }

        # Final receive/cleanup
        if ($jobs) {
            $more = $jobs | Receive-Job -Wait -AutoRemoveJob
            if ($more) { $more | Write-Output }
        }

        if ($ShowProgress) {
            Write-Progress -Id 1 -Activity 'Querying print servers' -Completed
        }
    }
}