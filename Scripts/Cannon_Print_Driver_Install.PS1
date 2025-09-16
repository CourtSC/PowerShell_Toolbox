function Install-RemotePrintDriver {
    <#
    .SYNOPSIS
    Installs a single, approved print driver on one or more remote computers. No other driver is allowed.

    .DESCRIPTION
    This command is hard-locked to install the **HP Universal Printing PCL 6 (v7.0.0)** driver, from the exact
    INF **hpcu250u.inf**, which must already be present/staged on the local machine. The function copies the local
    driver’s directory to each remote machine (C:\Temp\HPUPD) and uses pnputil to add that INF, then registers the
    driver via Add-PrinterDriver.

    You can target computers via -ComputerName (with optional -Credential) or pass existing -Session objects.
    Only sessions created by this function are removed on completion.

    The operation honors -WhatIf / -Confirm.

    .INPUTS
    System.String, Microsoft.PowerShell.Commands.PSSession

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    Properties: ComputerName, DriverName, Installed, Version, InfPath

    .PARAMETER ComputerName
    One or more computer names to install the driver on. Accepts pipeline input.

    .PARAMETER Session
    One or more existing PSSessions to use. If specified, -Credential is ignored.

    .PARAMETER Credential
    Credentials for remoting when using -ComputerName. (Ignored with -Session.)

    .PARAMETER ThrottleLimit
    Max concurrent calls when using -ComputerName. Default 16.

    .EXAMPLE
    PS> 'PC01','PC02' | Install-RemotePrintDriver -Credential (Get-Credential) -Verbose

    .EXAMPLE
    PS> $s = New-PSSession -ComputerName 'PC01','PC02'
    PS> Install-RemotePrintDriver -Session $s -Confirm

    .NOTES
    - Approved driver only: Name = "HP Universal Printing PCL 6 (v7.0.0)", INF = "hpcu250u.inf".
    - Requires the driver to be installed/staged on the local machine first.
    - Requires admin rights on target machines; uses pnputil and Add-PrinterDriver remotely.
    - PowerShell 7+ recommended.
    .LINK
    about_Remote
    about_CommonParameters
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'ComputerName')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(ParameterSetName = 'ComputerName', Position = 0, Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(ParameterSetName = 'Session', Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,

        [Parameter(ParameterSetName = 'ComputerName')]
        [pscredential]$Credential,

        [Parameter(ParameterSetName = 'ComputerName')]
        [ValidateRange(1, 128)]
        [int]$ThrottleLimit = 16
    )

    begin {
        # ==== HARD-LOCKED DRIVER METADATA (do not change to install anything else) ====
        $DriverName = 'Canon iR-ADV 8285/8295 UFR II'
        $ExpectedInfFile = 'cnlb0ua64.inf'
        $RemoteRoot = 'C:\Temp\Canon'

        # Validate the driver exists locally and locate its root folder
        try {
            $localDriver = Get-PrinterDriver -Name $DriverName -ErrorAction Stop
            $localInf = $localDriver.InfPath
            if (-not $localInf) { throw 'Local driver InfPath not found.' }
            $localRoot = Split-Path -Path $localInf -Parent
            if (-not (Test-Path -LiteralPath $localRoot)) { throw "Local driver folder not found at '$localRoot'." }
            # Check expected INF is present in the local tree
            $expectedLocalInf = Join-Path -Path $localRoot -ChildPath $ExpectedInfFile
            if (-not (Test-Path -LiteralPath $expectedLocalInf)) {
                throw "Expected INF '$ExpectedInfFile' not found under '$localRoot'."
            }
            Write-Verbose "Local driver validated: Name='$DriverName', Root='$localRoot', INF='$ExpectedInfFile'."
        } catch {
            Write-Error "Driver validation failed locally: $($_.Exception.Message)"
            return
        }

        # Build the remote scriptblock - installs ONLY the expected INF for the hard-locked driver
        $installScript = {
            param(
                [string]$DriverName,
                [string]$RemoteRoot,
                [string]$ExpectedInfFile
            )
            $ErrorActionPreference = 'Stop'

            # Ensure spooler is running
            if ((Get-Service -Name Spooler).Status -ne 'Running') {
                Start-Service -Name Spooler
            }

            # Validate remote INF exists where we expect it
            $infPath = Join-Path -Path $RemoteRoot -ChildPath $ExpectedInfFile
            if (-not (Test-Path -LiteralPath $infPath)) {
                throw "Expected INF '$ExpectedInfFile' not found at '$infPath'."
            }

            # Stage driver package (idempotent if already present)
            $pnputil = Join-Path $env:SystemRoot 'System32\pnputil.exe'
            if (-not (Test-Path -LiteralPath $pnputil)) {
                throw "pnputil not found at '$pnputil'."
            }

            # Add the exact INF, then register the driver by the exact name
            & $pnputil /add-driver $infPath /install | Out-String | Write-Verbose

            # Add-PrinterDriver is idempotent; if present, it's a no-op
            if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
                Add-PrinterDriver -Name $DriverName
            }

            # Return verification details
            $drv = Get-PrinterDriver -Name $DriverName -ErrorAction Stop
            [pscustomobject]@{
                ComputerName = $env:COMPUTERNAME
                DriverName   = $drv.Name
                Installed    = $true
                Version      = $drv.Version
                InfPath      = $drv.InfPath
            }
        }

        # Track sessions we create so we can clean up only those
        $ownedSessions = New-Object System.Collections.Generic.List[System.Management.Automation.Runspaces.PSSession]
        $targets = New-Object System.Collections.Generic.List[System.Management.Automation.Runspaces.PSSession]
    }

    process {
        # Normalize Session set
        if ($PSCmdlet.ParameterSetName -eq 'Session') {
            foreach ($s in $Session) { $targets.Add($s) }
        } else {
            # Create sessions for each ComputerName
            foreach ($cn in $ComputerName) {
                if ($PSCmdlet.ShouldProcess($cn, 'Create PSSession')) {
                    try {
                        $s = New-PSSession -ComputerName $cn -Credential $Credential -ErrorAction Stop
                        $ownedSessions.Add($s)
                        $targets.Add($s)
                        Write-Verbose "Session established: $cn"
                    } catch {
                        Write-Error "Failed to create session to '$cn': $($_.Exception.Message)"
                    }
                }
            }
        }

        if ($targets.Count -eq 0) { return }

        # Ensure remote folder exists, then copy files
        foreach ($s in $targets) {
            $cn = $s.ComputerName
            if ($PSCmdlet.ShouldProcess($cn, "Prepare folder '$RemoteRoot'")) {
                try {
                    Invoke-Command -Session $s -ErrorAction Stop -ScriptBlock {
                        param($RemoteRoot)
                        New-Item -Path $RemoteRoot -ItemType Directory -Force | Out-Null
                    } -ArgumentList $RemoteRoot
                } catch {
                    Write-Error "Failed to prepare folder on '$cn': $($_.Exception.Message)"
                    continue
                }
            }

            if ($PSCmdlet.ShouldProcess($cn, "Copy driver files to '$RemoteRoot'")) {
                try {
                    $source = Join-Path $localRoot '*'
                    Copy-Item -Path $source -Destination $RemoteRoot -ToSession $s -Recurse -Force | Out-Null                    # Confirm expected INF exists remotely after copy (quick sanity check)
                    $expectedRemoteInf = Join-Path -Path $RemoteRoot -ChildPath "cnlb0ua64.inf_amd64_ee7554a41c90e13c\$ExpectedInfFile"
                    $present = Invoke-Command -Session $s -ErrorAction Stop -ScriptBlock {
                        param($root, $file) 
                        $path = (Get-ChildItem -Path $root -Recurse -Filter $file).PSPath
                        Test-Path -Path $path
                    } -ArgumentList $RemoteRoot, $ExpectedInfFile
                    if (-not $present) {
                        throw "Post-copy validation failed: '$ExpectedInfFile' not present at '$expectedRemoteInf'."
                    }
                } catch {
                    Write-Error "File copy failed to '$cn': $($_.Exception.Message)"
                    continue
                }
            }

            # Install on this target
            if ($PSCmdlet.ShouldProcess($cn, "Install driver '$DriverName' using '$ExpectedInfFile'")) {
                try {
                    Invoke-Command -Session $s -ErrorAction Stop -ScriptBlock $installScript -ArgumentList $DriverName, $RemoteRoot, $ExpectedInfFile
                } catch {
                    Write-Error "Installation failed on '$cn': $($_.Exception.Message)"
                    continue
                }
            }
        }
    }

    end {
        # Clean up only sessions we created
        foreach ($s in $ownedSessions) {
            if ($PSCmdlet.ShouldProcess($s.ComputerName, 'Remove owned session')) {
                try { Remove-PSSession -Session $s -ErrorAction Stop } catch { Write-Verbose "Session cleanup warning: $($_.Exception.Message)" }
            }
        }
    }
}