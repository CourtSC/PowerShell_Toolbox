function Get-Cred {
        Unlock-SecretStore -Password (Read-Host -AsSecureString 'Enter vault password')
        $global:cred = Get-Secret -Name 'cred-a'
        return $global:cred
}

if (-not $global:cred) { Get-Cred }

# Use Type Accelerator to allow PSSession type.
[PowerShell].Assembly.GetType('System.Management.Automation.TypeAccelerators')::add('PSSession', 'System.Management.Automation.Runspaces.PSSession')

$mhdPrintServers = ('ADCVPRNMHDMS001', `
                'ADCVPRNMHDMS002', `
                'ADCVPRNMHDMS003', `
                'ADCVPRNMHDMS004', `
                'ADCVPRNMHDMS005', `
                'ADCVPRNMHDMS006'
)

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

if (-not $mhdPrintServers -and -not $fhdPrintServers) {}
