function Get-MHDPrinters {
    <#
    .SYNOPSIS
    Gets printer information from a set of MHD print servers.

    .DESCRIPTION
    Queries one or more print servers for one or more specified printer names.
    By default, queries all six MHD print servers (ADCVPRNMHDMS001–006).

    Returns structured objects with server name and printer details.
    Errors are surfaced but do not stop enumeration.

    .PARAMETER Printers
    One or more printer names to look up. Accepts pipeline input.

    .PARAMETER Servers
    One or more print server names to query. Defaults to the six MHD servers.

    .EXAMPLE
    PS> Get-MHDPrinter -Printers 'HP123','CanonX'

    .EXAMPLE
    PS> 'HP123' | Get-MHDPrinter

    .NOTES
    Requires the PrintManagement module (`Get-Printer`).
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
        [string[]]$Servers = ('ADCVPRNMHDMS001', `
                'ADCVPRNMHDMS002', `
                'ADCVPRNMHDMS003', `
                'ADCVPRNMHDMS004', `
                'ADCVPRNMHDMS005', `
                'ADCVPRNMHDMS006', `
                'ADCVPRNMHDMS020', `
                'ADCVPRNMHDMS021', `
                'ADCVPRNMHDMS022', `
                'ADCVPRNMHDMS023', `
                'ADCVPRNMHDMS024', `
                'ADCVPRNMHDMS030', `
                'ADCVPRNMHDMS031', `
                'ADCVPRNMHDMS040', `
                'ADCVPRNMHDMS041', `
                'ADCVPRNMHDMS050', `
                'ADCVPRNMHDMS051', `
                'ADCVPRNMHDMS052', `
                'ADCVPRNMHDMS060', `
                'ADCVPRNMHDMS061', `
                'ADCVPRNMHDMS062', `
                'ADCVPRNMHDMS070', `
                'ADCVPRNMHDMS071', `
                'ADCVPRNMHDMS080', `
                'ADCVPRNMHDMS081')
    )
    begin {
        Import-Module ActiveDirectory
        $total = $printers.count
        $count = 0
        function Get-Progress {
            param (
                $Total, $Count = 0, $Message
            )
            $progress = [Math]::Round(($Count / $Total) * 100, 2)
            Write-Progress -Activity $Message -Status "$progress% complete" -PercentComplete $progress
        }
    }

    process {
        if (!$PSBoundParameters.ContainsKey('Printers')) {
            Write-Verbose 'No printers provided with function call. Looking up all printers.'
            $allPrinters = $Servers | ForEach-Object -Parallel {
                Get-Printer -ComputerName $_
            }
            $total = $allPrinters.count
            $output = foreach ($p in $allPrinters) {
                $count++
                Get-Progress -Total $total -Count $count -Message "Checking $($p.ComputerName) for $($p.Name)..."
                $port = Get-PrinterPort -ComputerName $p.ComputerName -Name $p.PortName -ErrorAction Stop
                [pscustomobject]@{
                    Server          = $p.ComputerName
                    Name            = $p.Name
                    Status          = $p.PrinterStatus
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
                $filterString = "*-PRN-$($p.Name)*"
                $groups = Get-ADGroup -Filter { Name -like $filterString } -ErrorAction Stop
                foreach ($group in $groups) {
                    if (($group.Name -like '*-DEF') -and ($group.Name -like '*-PRN-*')) {
                        $output | Add-Member -NotePropertyName 'DefaultGroup' -NotePropertyValue $group.Name -Force
                    } elseif ($group.Name -like '*-PRN-*') {
                        $output | Add-Member -NotePropertyName 'StandardGroup' -NotePropertyValue $group.Name -Force
                    }
                }
            }
            return $output
        } else {
            $output = foreach ($printer in $Printers) {
                foreach ($srv in $Servers) {
                    Get-Progress -Total $total -Count $count -Message "Checking $($srv) for $($printer)..."
                    try {
                        $p = Get-Printer -ComputerName $srv -Name $printer -ErrorAction Stop
                        if ($p) {
                            $port = Get-PrinterPort -ComputerName $srv -Name $p.PortName -ErrorAction Stop
                            $filterString = "*-PRN-$($p.Name)*"
                            $groups = Get-ADGroup -Filter { Name -like $filterString } -ErrorAction Stop
                            foreach ($group in $groups) {
                                if (($group.Name -like '*-DEF') -and ($group.Name -like '*-PRN-*')) {
                                    $defGroup = $group.Name
                                } elseif ($group.Name -like '*-PRN-*') {
                                    $stdGroup = $group.Name
                                }
                            }
                        }
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
                    } catch {
                        Write-Verbose "Printer '$printer' not found on $srv. ($($_.Exception.Message))"
                    }
                }
                $count++
            }
            return $output
        }
    }
}