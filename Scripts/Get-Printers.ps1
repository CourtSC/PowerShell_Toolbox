$output = @()
$mhdPrintServers = @(
    'ADCVPRNMHDMS001', 'ADCVPRNMHDMS002', 'ADCVPRNMHDMS003', 'ADCVPRNMHDMS004', 'ADCVPRNMHDMS005', 'ADCVPRNMHDMS006',
    'ADCVPRNMHDMS020', 'ADCVPRNMHDMS021', 'ADCVPRNMHDMS022', 'ADCVPRNMHDMS023', 'ADCVPRNMHDMS024',
    'ADCVPRNMHDMS030', 'ADCVPRNMHDMS031',
    'ADCVPRNMHDMS040', 'ADCVPRNMHDMS041',
    'ADCVPRNMHDMS050', 'ADCVPRNMHDMS051', 'ADCVPRNMHDMS052',
    'ADCVPRNMHDMS060', 'ADCVPRNMHDMS061', 'ADCVPRNMHDMS062',
    'ADCVPRNMHDMS070', 'ADCVPRNMHDMS071',
    'ADCVPRNMHDMS080', 'ADCVPRNMHDMS081'
)
$siteCodes = ( 'ALT', 'APK', 'CEL', 'CEN', 'COU', 'FMG', 'EOR', 'KIS', 'ORL', 'WGA', 'WPK' )
foreach ($srv in $mhdPrintServers) {
    $printers = Get-Printer -ComputerName $srv | Select-Object Name, ComputerName, DriverName, ShareName, Shared, Published, PortName, Location, Comment
    $ports = Get-PrinterPort -ComputerName $srv
    foreach ($p in $printers) {
        $groups = @()
        $hostAddress = $null
        $stdGroup = $null
        $stdInfo = $null
        $defGroup = $null
        $defInfo = $null
        $hostAddress = $ports | Where-Object {
            ($_.PrinterHostAddress) -and
            ($p.PortName -eq $_.Name)
        } | Select-Object -ExpandProperty PrinterHostAddress
        foreach ($site in $siteCodes) {
            $groupString = "$site-PRN-$($p.Name)"
            try { $groups += Get-ADGroup -Identity "$groupString-DEF" -Properties info } catch {}
            try { $groups += Get-ADGroup -Identity $groupString -Properties info } catch {}
        }
        foreach ($g in $groups) {
            if ($g.info.Substring(2, 15) -eq $p.ComputerName) {
                if ($g.Name -like '*-DEF') { 
                    $defGroup = $g.Name
                    $defInfo = $g.info
                } elseif ($g.Name -like '*-PRN-*') {
                    $stdGroup = $g.Name
                    $stdInfo = $g.info
                }
            }
        }
        if ($hostAddress) {
            $output += [pscustomobject]@{
                Name          = $p.Name
                ComputerName  = $p.ComputerName
                DriverName    = $p.DriverName
                ShareName     = $p.ShareName
                Shared        = $p.Shared
                Published     = $p.Published
                PortName      = $p.PortName
                Location      = $p.Location
                Comment       = $p.Comment
                HostAddress   = $hostAddress
                StandardGroup = $stdGroup
                StandardNotes = $stdInfo
                DefaultGroup  = $defGroup
                DefaultNotes  = $defInfo
            }
        }
    }
}

$outputJSON = $output | ConvertTo-Json

return $outputJSON