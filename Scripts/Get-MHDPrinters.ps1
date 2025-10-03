<#
.SYNOPSIS
Gets printer details from one or more print servers and enriches them with
Active Directory group info and TCP/IP port metadata.

.DESCRIPTION
Get-MHDPrinters queries the specified print servers for printers (optionally
filtering by printer name). For each printer, it looks up related AD groups
by name pattern "*-PRN-<PrinterName>*" and fetches the associated TCP/IP port
to report host address, SNMP status, and description.

This command uses PowerShell 7+ parallel execution to query servers. Use
-Verbose for detailed per-server/per-printer trace output. Failures are logged
via Write-Error/Write-Warning but do not stop the entire command; missing
fields are returned as $null where appropriate.

.PARAMETER Printers
Optional list of printer names to include. When specified, only printers
whose Name is in this list are returned.

.PARAMETER Servers
List of print servers to query. Defaults to the known MHD servers.

.PARAMETER ThrottleLimit
Maximum number of parallel runspaces used when querying servers. Default: 8.

.PARAMETER ADSearchBase
Optional distinguished name (DN) to limit the AD group search scope.
If not specified, the default AD search scope is used.

.INPUTS
System.String
You can pipe printer names (by property name or value) to -Printers.

.OUTPUTS
System.Management.Automation.PSObject
One object per printer with server, driver, port, AD group, and status fields.

.EXAMPLE
PS> Get-MHDPrinters -Verbose
Queries all default servers and prints verbose progress and any warnings.

.EXAMPLE
PS> 'PRN-EDU-1','PRN-EDU-2' | Get-MHDPrinters -Servers ADCVPRNMHDMS001 -Verbose
Queries a single server for the two specified printers and logs details.

.NOTES
Requires modules: PrintManagement, ActiveDirectory
PowerShell 7+ is required for ForEach-Object -Parallel.

.LINK
Get-Printer
Get-PrinterPort
Get-ADGroup
ForEach-Object
#>
function Get-MHDPrinters {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]] $Printers,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]] $Servers = @(
            'ADCVPRNMHDMS001', 'ADCVPRNMHDMS002', 'ADCVPRNMHDMS003', 'ADCVPRNMHDMS004', 'ADCVPRNMHDMS005', 'ADCVPRNMHDMS006',
            'ADCVPRNMHDMS020', 'ADCVPRNMHDMS021', 'ADCVPRNMHDMS022', 'ADCVPRNMHDMS023', 'ADCVPRNMHDMS024',
            'ADCVPRNMHDMS030', 'ADCVPRNMHDMS031',
            'ADCVPRNMHDMS040', 'ADCVPRNMHDMS041',
            'ADCVPRNMHDMS050', 'ADCVPRNMHDMS051', 'ADCVPRNMHDMS052',
            'ADCVPRNMHDMS060', 'ADCVPRNMHDMS061', 'ADCVPRNMHDMS062',
            'ADCVPRNMHDMS070', 'ADCVPRNMHDMS071',
            'ADCVPRNMHDMS080', 'ADCVPRNMHDMS081'
        ),

        [Parameter()]
        [ValidateRange(1, 64)]
        [int] $ThrottleLimit = 8,

        [Parameter()]
        [string] $ADSearchBase
    )

    begin {
        try {
            Write-Verbose 'Importing required modules in main runspace...'
            Import-Module PrintManagement -ErrorAction Stop
            Import-Module ActiveDirectory -ErrorAction Stop
        } catch {
            Write-Error -Category ResourceUnavailable -ErrorId 'ModuleImportFailed' `
                -Message "Failed to import required modules (PrintManagement/ActiveDirectory). $_"
            return
        }

        if ($Printers) {
            Write-Verbose ('Printer filter active: {0}' -f ($Printers -join ', '))
        }
        Write-Verbose ('Servers to query: {0}' -f ($Servers -join ', '))
    }

    process {
        # 1) PARALLEL PHASE: emit basic printer objects; pipeline will aggregate them
        $printerObjects = $Servers | ForEach-Object -Parallel {
            # Honor caller's -Verbose inside parallel runspaces
            $VerbosePreference = $using:VerbosePreference

            try {
                Import-Module PrintManagement -ErrorAction Stop
            } catch {
                Write-Error -Category ResourceUnavailable -ErrorId 'ModuleImportInParallelFailed' `
                    -Message "[$_] Failed to import PrintManagement in parallel runspace. $_"
                return
            }

            $server = $_
            Write-Verbose "[$server] Starting printer query..."

            try {
                $printersFromServer = Get-Printer -ComputerName $server -ErrorAction Stop

                if ($using:Printers) {
                    $printersFromServer = $printersFromServer | Where-Object {
                        $using:Printers -contains $_.Name
                    }
                    Write-Verbose "[$server] Filtered printers by name; remaining: $(@($printersFromServer).Count)"
                } else {
                    Write-Verbose "[$server] Retrieved $(@($printersFromServer).Count) printers"
                }

                foreach ($p in $printersFromServer) {
                    # EMIT to pipeline (no shared collection)
                    [pscustomobject]@{
                        ComputerName  = $server
                        Name          = $p.Name
                        DriverName    = $p.DriverName
                        PortName      = $p.PortName
                        Shared        = $p.Shared
                        Published     = $p.Published
                        Comment       = $p.Comment
                        Location      = $p.Location
                        PrinterStatus = $p.PrinterStatus
                    }
                }
            } catch {
                Write-Error -Category ResourceUnavailable -ErrorId 'PrinterQueryFailed' `
                    -Message "[$server] Failed to query printers. $_"
            } finally {
                Write-Verbose "[$server] Finished printer query."
            }
        } -ThrottleLimit $ThrottleLimit

        # 2) ENRICHMENT PHASE (sequential): AD groups + port metadata
        foreach ($p in $printerObjects) {
            $filterString = "*-PRN-$($p.Name)*"
            $stdGroup = $null
            $defGroup = $null

            # AD group lookup
            try {
                $adParams = @{
                    Filter      = "Name -like '$filterString'"
                    ErrorAction = 'Stop'
                }
                if ($ADSearchBase) { $adParams['SearchBase'] = $ADSearchBase }

                Write-Verbose ("[AD] Searching groups with filter '{0}'{1}..." -f $adParams.Filter, $(if ($ADSearchBase) { " in '$ADSearchBase'" } else { '' }))
                $groups = Get-ADGroup @adParams

                $groupCount = @($groups).Count
                Write-Verbose "[AD] Found $groupCount group(s) for printer '$($p.Name)'."

                foreach ($g in $groups) {
                    if ($g.Name -like '*-PRN-*' -and $g.Name -like '*-DEF') { $defGroup = $g.Name }
                    elseif ($g.Name -like '*-PRN-*') { $stdGroup = $g.Name }
                }
            } catch {
                Write-Warning ("[AD] Failed to search groups for printer '{0}'. {1}" -f $p.Name, $_.Exception.Message)
            }

            # Port lookup
            $port = $null
            try {
                Write-Verbose ("[Port] Querying port '{0}' on '{1}'..." -f $p.PortName, $p.ComputerName)
                $port = Get-PrinterPort -Name $p.PortName -ComputerName $p.ComputerName -ErrorAction Stop
            } catch {
                Write-Warning ("[Port] Failed to get port '{0}' on '{1}'. {2}" -f $p.PortName, $p.ComputerName, $_.Exception.Message)
            }

            # Emit final enriched object
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
    }
}
