function Get-Cred {
        Unlock-SecretStore -Password (Read-Host -AsSecureString 'Enter vault password')
        $global:cred = Get-Secret -Name 'cred-a'
        return $global:cred
}

if (-not $global:cred) { Get-Cred }

Import-Module -Name ITSSFunctions

# Use Type Accelerator to allow PSSession type.
[PowerShell].Assembly.GetType('System.Management.Automation.TypeAccelerators')::add('PSSession', 'System.Management.Automation.Runspaces.PSSession')

$mhdPrintServers = ('ADCVPRNMHDMS001', `
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

$fhdPrintServers = ('FHOSVMWPRN001', `
                'FHOSVMWPRN002', `
                'FHOSVMWPRN003', `
                'FHOSVMWPRN004', `
                'FHOSVMWPRN005', `
                'FHOSVMWPRN006', `
                'FHOVPRNA001', `
                'FHOVPRNB001', `
                'FHOVPRNC001', `
                'FHOVPRND001', `
                'FHOVPRNE001', `
                'FHOVPRNF001')

$dhcpSrvs = @('Orlvnetfhdms001.flhosp.net', 'Fhovnetfhdms001.flhosp.net')

if (-not $mhdPrintServers -and -not $fhdPrintServers -and -not $dhcpSrvs) {}
