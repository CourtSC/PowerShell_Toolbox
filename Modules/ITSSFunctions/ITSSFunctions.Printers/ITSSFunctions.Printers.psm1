function Get-MHDPrinters {
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

            $config = $null
            $config = Get-PrintConfiguration -ComputerName $p.ComputerName -PrinterName $p.Name

            # Emit final enriched object
            [pscustomobject]@{
                Server          = $p.ComputerName
                Name            = $p.Name
                Status          = $p.PrinterStatus
                StandardGroup   = $stdGroup
                DefaultGroup    = $defGroup
                DriverName      = $p.DriverName
                Color           = $config.Color
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

function Get-PrintersWithErrors {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Domain = 'multihosp.net',

        [Parameter()]
        [string[]]$Servers,

        [Parameter()]
        [switch]$Remove,

        [Parameter()]
        [switch]$Nuke
    )

    if ($Nuke) { $Remove = $true }

    if (($Domain -eq 'multihosp.net') -and (-not $Servers)) {
        $Servers = @(
            'ADCVPRNMHDMS001', 'ADCVPRNMHDMS002', 'ADCVPRNMHDMS003', 'ADCVPRNMHDMS004',
            'ADCVPRNMHDMS005', 'ADCVPRNMHDMS006'
        )
    }

    $printers = foreach ($server in $Servers) {
        try {
            Get-Printer -ComputerName $server -ErrorAction Stop | Where-Object { ($_.PrinterStatus -eq 'Error') -or (($_.PrinterStatus -eq 'Offline')) }
        } catch {
            Write-Warning "Failed to get printers from $($server): $_"
        }
    }

    if ($Remove) {
        # Process print jobs and collect results
        $results = $printers | ForEach-Object -Parallel {
            $printer = $_
            $summary = @()

            try {
                if ($Nuke) {
                    $jobs = Get-PrintJob -PrinterObject $printer
                } else {
                    $jobs = Get-PrintJob -PrinterObject $printer | Where-Object { $_.JobStatus -like '*error*' }
                }

                foreach ($job in $jobs) {
                    $status = 'Success'
                    $message = ''

                    try {
                        Remove-PrintJob -InputObject $job -ErrorAction Stop
                    } catch {
                        $status = 'Failed'
                        $message = $_.Exception.Message

                        if ($message -match 'Access was denied to the specified resource') {
                            Write-Warning "Access denied on [$($printer.ComputerName)] for [$($printer.Name)]"
                        } else {
                            Write-Warning "Failed to remove job on [$($printer.ComputerName)]: $message"
                        }
                    }

                    $summary += [PSCustomObject]@{
                        Server   = $printer.ComputerName
                        Printer  = $printer.Name
                        JobId    = $job.ID
                        Document = $job.DocumentName
                        Status   = $status
                        Message  = $message
                    }
                }
            } catch {
                $summary += [PSCustomObject]@{
                    Server   = $printer.ComputerName
                    Printer  = $printer.Name
                    JobId    = $null
                    Document = $null
                    Status   = 'Failed'
                    Message  = "Failed to retrieve print jobs: $($_.Exception.Message)"
                }
            }

            return $summary
        } -ThrottleLimit 5
    } else { $results = $printers }

    # Flatten and return all job results
    return $results | ForEach-Object { $_ } | Sort-Object Server, Printer, Status
    
}

function Remove-PrinterPortSNMP {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter()]
        [string[]] $Servers = @(
            'ADCVPRNMHDMS001',
            'ADCVPRNMHDMS002',
            'ADCVPRNMHDMS003',
            'ADCVPRNMHDMS004',
            'ADCVPRNMHDMS005',
            'ADCVPRNMHDMS006'
        ),
        [string] $LogPath
    )

    begin {
        $script:LogToFile = { param($Level, $Message)
            if ($PSBoundParameters.ContainsKey('LogPath')) {
                $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
                Add-Content -Path $LogPath -Value $line
            }
        }

        function Write-LogInfo { param($m) Write-Information $m -InformationAction Continue; & $script:LogToFile 'INFO' $m }
        function Write-LogWarn { param($m) Write-Warning $m; & $script:LogToFile 'WARN' $m }
        function Write-LogVerbose { param($m) Write-Verbose $m; & $script:LogToFile 'VERBOSE' $m }
        function Write-LogError {
            param($m, $e)
            $detail = if ($e) { '{0} | {1} | {2}' -f $e.FullyQualifiedErrorId, $e.CategoryInfo, $e.Exception.Message } else { '' }
            Write-Error -Message "$m $detail"
            & $script:LogToFile 'ERROR' "$m $detail"
        }

        if ($PSBoundParameters.ContainsKey('LogPath')) {
            $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
            Add-Content -Path $LogPath -Value ('=' * 80)
            Add-Content -Path $LogPath -Value ('Run started {0}' -f (Get-Date))
        }

        Write-LogInfo "Starting Remove-PrinterPortSNMP on $($Servers.Count) server(s)."
    }

    process {
        foreach ($srv in $Servers) {
            Write-LogInfo "---- Server: $srv ----"

            try {
                $ports = Get-PrinterPort -ComputerName $srv -ErrorAction Stop | Where-Object { $_.SNMPEnabled }

                if (-not $ports) {
                    Write-LogVerbose "No SNMP-enabled ports found on $srv."
                    continue
                }

                Write-LogInfo ('Found {0} SNMP-enabled port(s) on {1}.' -f $ports.Count, $srv)

                foreach ($port in $ports) {
                    $portSummary = "Port '{0}' ({1}:{2})" -f $port.Name, $port.PrinterHostAddress, $port.PortNumber
                    Write-LogInfo "Processing $portSummary on $srv."

                    try {
                        if ($PSCmdlet.ShouldProcess("$srv | $portSummary", 'Replace port to disable SNMP')) {

                            $tempName = 'temp'
                            $existingTemp = Get-PrinterPort -ComputerName $srv -Name $tempName -ErrorAction SilentlyContinue
                            if (-not $existingTemp) {
                                Add-PrinterPort -ComputerName $srv -Name $tempName -PrinterHostAddress $port.PrinterHostAddress -PortNumber $port.PortNumber -ErrorAction Stop
                                Write-LogVerbose "Created temp port on $srv."
                            } else {
                                Write-LogVerbose "Temp port already exists on $srv; reusing."
                            }

                            $printersOnPort = Get-Printer -ComputerName $srv -ErrorAction Stop | Where-Object { $_.PortName -eq $port.Name }
                            if ($printersOnPort) {
                                foreach ($printer in $printersOnPort) {
                                    Set-Printer -InputObject $printer -PortName $tempName -ErrorAction Stop
                                    Write-LogInfo "Repointed printer '{0}' to temp." -f $printer.Name
                                }
                            } else {
                                Write-LogVerbose "No printers currently using $($port.Name) on $srv."
                            }

                            Remove-PrinterPort -InputObject $port -ErrorAction Stop
                            Write-LogInfo "Removed SNMP-enabled port $($port.Name) on $srv."

                            Add-PrinterPort -ComputerName $srv -Name $port.Name -PrinterHostAddress $port.PrinterHostAddress -PortNumber $port.PortNumber -ErrorAction Stop
                            Write-LogInfo "Recreated port $($port.Name) on $srv (SNMP disabled)."

                            if ($printersOnPort) {
                                foreach ($printer in $printersOnPort) {
                                    Set-Printer -InputObject $printer -PortName $port.Name -ErrorAction Stop
                                    Write-LogInfo "Repointed printer '{0}' back to {1}." -f $printer.Name, $port.Name
                                }
                            }

                            $printersOnTemp = Get-Printer -ComputerName $srv -ErrorAction Stop | Where-Object { $_.PortName -eq $tempName }
                            if (-not $printersOnTemp) {
                                Remove-PrinterPort -ComputerName $srv -Name $tempName -ErrorAction SilentlyContinue
                                Write-LogVerbose "Removed temp port from $srv."
                            } else {
                                Write-LogWarn "Temp port still in use on $srv; skipping removal."
                            }

                            Write-LogInfo "Completed $portSummary on $srv."
                        }
                    } catch {
                        Write-LogError "Failed while processing $portSummary on $srv." $_
                        continue
                    }
                }
            } catch {
                Write-LogError "Failed to enumerate ports or printers on $srv." $_
                continue
            }
        }
    }

    end {
        Write-LogInfo 'Remove-PrinterPortSNMP run complete.'
        if ($PSBoundParameters.ContainsKey('LogPath')) {
            Add-Content -Path $LogPath -Value ('Run ended {0}' -f (Get-Date))
            Add-Content -Path $LogPath -Value ('=' * 80)
        }
    }
}

function Install-RemotePrintDriver {
    <#
    .SYNOPSIS
    Installs a single, approved print driver on one or more remote computers. No other driver is allowed.

    .DESCRIPTION
    This command is hard-locked to install the **HP Universal Printing PCL 6 (v7.0.0)** driver, from the exact
    INF **hpcu250u.inf**, which must already be present/staged on the local machine. The function copies the local
    driver’s directory to each remote machine (C:\Temp\HPUPD) and uses pnputil to add that INF, then registers the
    driver via Add-PrinterDriver.

    In addition, to support larger remoting payloads, it ensures **WSMan MaxEnvelopeSizekb = 4096** on the **local**
    computer (before any remoting starts) and on **each remote** computer as soon as a session is established. If the
    current value is already >= 4096 it is left unchanged.

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
    - Ensures WSMan MaxEnvelopeSizekb >= 4096 locally and on each remote. Administrative rights are required.
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
        $DriverName = 'HP Universal Printing PCL 6 (v7.0.0)'
        $ExpectedInfFile = 'hpcu250u.inf'
        $RemoteRoot = 'C:\Temp\HPUPD'
        $DesiredEnvelopeKB = 4096

        # Helper scriptblock to ensure WSMan MaxEnvelopeSizekb on remote
        $ensureEnvelopeScript = {
            param([int]$DesiredKB)
            $ErrorActionPreference = 'Stop'
            try {
                $item = Get-Item -Path WSMan:\localhost\MaxEnvelopeSizekb
                $current = [int]$item.Value
            } catch {
                throw "Unable to read WSMan MaxEnvelopeSizekb: $($_.Exception.Message)"
            }

            if ($current -lt $DesiredKB) {
                Set-Item -Path WSMan:\localhost\MaxEnvelopeSizekb -Value $DesiredKB -Force | Out-Null
            }

            # Return the effective setting for visibility
            [pscustomobject]@{
                ComputerName      = $env:COMPUTERNAME
                MaxEnvelopeSizekb = (Get-Item WSMan:\localhost\MaxEnvelopeSizekb).Value
                Changed           = ($current -lt $DesiredKB)
            }
        }

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

        # Ensure LOCAL WSMan MaxEnvelopeSizekb >= 4096 BEFORE remoting
        try {
            $localItem = Get-Item -Path WSMan:\localhost\MaxEnvelopeSizekb -ErrorAction Stop
            $localCurrent = [int]$localItem.Value
            if ($localCurrent -lt $DesiredEnvelopeKB) {
                if ($PSCmdlet.ShouldProcess('localhost', "Set WSMan MaxEnvelopeSizekb to $DesiredEnvelopeKB")) {
                    Set-Item -Path WSMan:\localhost\MaxEnvelopeSizekb -Value $DesiredEnvelopeKB -Force
                    Write-Verbose "WSMan MaxEnvelopeSizekb on localhost changed from $localCurrent to $DesiredEnvelopeKB."
                }
            } else {
                Write-Verbose "WSMan MaxEnvelopeSizekb on localhost already $localCurrent (>= $DesiredEnvelopeKB)."
            }
        } catch {
            Write-Error "Failed to ensure local WSMan MaxEnvelopeSizekb: $($_.Exception.Message)"
            return
        }

        # Build the remote install script (installs ONLY the expected INF for the hard-locked driver)
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
        # Normalize Session set or create new sessions
        if ($PSCmdlet.ParameterSetName -eq 'Session') {
            foreach ($s in $Session) { $targets.Add($s) }
        } else {
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

        # Ensure REMOTE WSMan MaxEnvelopeSizekb >= 4096 on each target BEFORE copy/installation
        foreach ($s in $targets) {
            $cn = $s.ComputerName
            if ($PSCmdlet.ShouldProcess($cn, "Ensure WSMan MaxEnvelopeSizekb >= $DesiredEnvelopeKB")) {
                try {
                    $result = Invoke-Command -Session $s -ErrorAction Stop -ScriptBlock $ensureEnvelopeScript -ArgumentList $DesiredEnvelopeKB
                    Write-Verbose ('WSMan MaxEnvelopeSizekb on {0}: {1} (Changed={2})' -f $result.ComputerName, $result.MaxEnvelopeSizekb, $result.Changed)
                } catch {
                    Write-Error "Failed to ensure WSMan MaxEnvelopeSizekb on '$cn': $($_.Exception.Message)"
                    continue
                }
            }
        }

        # Ensure remote folder exists, then copy files and install
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
                    Copy-Item -Path $source -Destination $RemoteRoot -ToSession $s -Recurse -Force | Out-Null

                    # Quick sanity check: confirm expected INF shows up somewhere under the root
                    $present = Invoke-Command -Session $s -ErrorAction Stop -ScriptBlock {
                        param($root, $file)
                        $found = Get-ChildItem -Path $root -Recurse -Filter $file -ErrorAction SilentlyContinue
                        [bool]$found
                    } -ArgumentList $RemoteRoot, $ExpectedInfFile

                    if (-not $present) {
                        throw "Post-copy validation failed: '$ExpectedInfFile' not found anywhere under '$RemoteRoot'."
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
                try { Remove-PSSession -Session $s -ErrorAction Stop }
                catch { Write-Verbose "Session cleanup warning: $($_.Exception.Message)" }
            }
        }
    }
}


function Set-PrinterPort {
    <#
    .SYNOPSIS
    Repairs a printer’s TCP/IP port configuration on a print server by using a temporary port.

    .DESCRIPTION
    This function fixes misconfigured TCP/IP ports for a given printer on a target print server. Because there is no
    PowerShell cmdlet to modify an existing port in place, it:
    1) Creates a unique temporary port with the desired host address,
    2) Repoints the printer to that temporary port,
    3) Removes the existing (incorrect) port if present,
    4) Recreates the correct port with the desired host address,
    5) Repoints the printer back to the correct port,
    6) Optionally cleans up the temporary port.

    The function supports -WhatIf/-Confirm and can run non-interactively with -Force to skip confirmation prompts.

    .PARAMETER Name
    The printer name on the target print server.

    .PARAMETER Server
    The print server (ComputerName) hosting the printer.

    .PARAMETER HostAddress
    The desired TCP/IP host address for the port (IP or DNS host).

    .PARAMETER Force
    Skips interactive confirmations (ShouldContinue) while still honoring -WhatIf/-Confirm.

    .PARAMETER NoCleanup
    Keeps the temporary port after the operation (useful for troubleshooting). By default the temp port is removed.

    .INPUTS
    System.String
    Accepts objects with Name/Server/HostAddress via property-based pipeline.

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    Emits: Server, PrinterName, FromPort, ToPort, TempPort, PortRecreated, Success, Message

    .EXAMPLES
    # Fix the port config for a single printer
    PS> Set-PrinterPort -Name 'HP-LJ-05' -Server 'PRINT01' -HostAddress '10.12.34.56' -Verbose

    # Run non-interactively (no prompts), still supports -WhatIf/-Confirm
    PS> Set-PrinterPort -Name 'HP-LJ-05' -Server 'PRINT01' -HostAddress '10.12.34.56' -Force

    # Pipe objects with the required properties
    PS> [pscustomobject]@{ Name='HP-A1'; Server='PRINT01'; HostAddress='prn-a1.contoso.com' } | Set-PrinterPort -Force
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('Printer', 'PrinterName')]
        [string]$Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('ComputerName')]
        [string]$Server,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('PrinterHostAddress', 'Address', 'Host')]
        [string]$HostAddress,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$NoCleanup
    )

    begin {
        function New-TempPortName {
            # Highly unlikely to collide; short GUID keeps names readable
            return ('zz-temp-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
        }
        # Canonical test scriptblocks
        $sbPortExists = {
            param($PortName, $Server)
            $null -ne (Get-PrinterPort -ComputerName $Server -Name $PortName -ErrorAction SilentlyContinue)
        }
        $sbGetPrinter = {
            param($Name, $Server)
            Get-Printer -Name $Name -ComputerName $Server -ErrorAction Stop
        }
    }

    process {
        $tempPortName = New-TempPortName
        $fromPort = $null
        $finalPort = $HostAddress
        $recreated = $false

        try {
            # Ensure printer exists and capture its current port
            $printer = & $sbGetPrinter $Name $Server
            $fromPort = $printer.PortName
        } catch {
            Write-Error "[$Server] Printer '$Name' not found or inaccessible: $($_.Exception.Message)"
            return
        }

        # 1) Ensure a temp port with the DESIRED host address exists
        if ($PSCmdlet.ShouldProcess($Server, "Create temporary port '$tempPortName' for $HostAddress")) {
            try {
                # Create temp with the target address; if name collides (very unlikely), generate a new one
                while (& $sbPortExists $tempPortName $Server) {
                    $tempPortName = New-TempPortName
                }
                Add-PrinterPort -ComputerName $Server -Name $tempPortName -PrinterHostAddress $HostAddress -ErrorAction Stop
                Write-Verbose "[$Server] Temp port '$tempPortName' created with host $HostAddress."
            } catch {
                Write-Error "[$Server] Failed creating temp port '$tempPortName' for host $($HostAddress): $($_.Exception.Message)"
                return
            }
        }

        # 2) Point the printer to the temp port (safe detour)
        if ($PSCmdlet.ShouldProcess($Server, "Repoint printer '$Name' to temp port '$tempPortName'")) {
            try {
                Set-Printer -Name $Name -ComputerName $Server -PortName $tempPortName -ErrorAction Stop
                Write-Verbose "[$Server] '$Name' temporarily moved to port '$tempPortName'."
            } catch {
                Write-Error "[$Server] Failed to assign temp port '$tempPortName' to '$Name': $($_.Exception.Message)"
                return
            }
        }

        # 3) Remove the existing (possibly incorrect) final port if it exists
        $finalExists = & $sbPortExists $finalPort $Server
        if ($finalExists) {
            $caption = "Remove port $finalPort on $Server?"
            $message = "Port $finalPort will be removed and recreated with the desired host address."
            if ($Force -or $PSCmdlet.ShouldContinue($message, $caption)) {
                if ($PSCmdlet.ShouldProcess($Server, "Remove port '$finalPort'")) {
                    try {
                        Remove-PrinterPort -ComputerName $Server -Name $finalPort -ErrorAction Stop
                        Write-Verbose "[$Server] Port '$finalPort' removed."
                        $finalExists = $false
                    } catch {
                        Write-Error "[$Server] Failed removing port '$finalPort': $($_.Exception.Message)"
                        # Try to roll back printer to its original port to leave system usable
                        try {
                            Set-Printer -Name $Name -ComputerName $Server -PortName $fromPort -ErrorAction Stop
                        } catch { }
                        return
                    }
                }
            } else {
                Write-Verbose "[$Server] Removal of port '$finalPort' cancelled by user."
            }
        }

        # 4) Create (or recreate) the final port with the desired address
        if (-not $finalExists) {
            $caption = "Create port $finalPort on $Server?"
            $message = "Port $finalPort will be created with host address $HostAddress."
            if ($Force -or $PSCmdlet.ShouldContinue($message, $caption)) {
                if ($PSCmdlet.ShouldProcess($Server, "Create port '$finalPort' for host $HostAddress")) {
                    try {
                        Add-PrinterPort -ComputerName $Server -Name $finalPort -PrinterHostAddress $HostAddress -ErrorAction Stop
                        $recreated = $true
                        Write-Verbose "[$Server] Port '$finalPort' created with host $HostAddress."
                    } catch {
                        Write-Error "[$Server] Failed creating port '$finalPort': $($_.Exception.Message)"
                        # Rollback printer to original port
                        try {
                            Set-Printer -Name $Name -ComputerName $Server -PortName $fromPort -ErrorAction Stop
                        } catch { }
                        return
                    }
                }
            } else {
                Write-Verbose "[$Server] Creation of port '$finalPort' cancelled by user."
            }
        }

        # 5) Repoint printer back to the final port
        if ($PSCmdlet.ShouldProcess($Server, "Repoint printer '$Name' to final port '$finalPort'")) {
            try {
                Set-Printer -Name $Name -ComputerName $Server -PortName $finalPort -ErrorAction Stop
                Write-Verbose "[$Server] '$Name' moved to final port '$finalPort'."
            } catch {
                Write-Error "[$Server] Failed to assign final port '$finalPort' to '$Name': $($_.Exception.Message)"
                return
            }
        }

        # 6) Cleanup temp port (unless requested to keep)
        if (-not $NoCleanup) {
            $caption = "Remove temporary port $tempPortName on $Server?"
            $message = 'Temp port will be removed.'
            if ($Force -or $PSCmdlet.ShouldContinue($message, $caption)) {
                if ($PSCmdlet.ShouldProcess($Server, "Remove temp port '$tempPortName'")) {
                    try {
                        Remove-PrinterPort -ComputerName $Server -Name $tempPortName -ErrorAction Stop
                        Write-Verbose "[$Server] Temp port '$tempPortName' removed."
                    } catch {
                        Write-Verbose "[$Server] Warning: temp port '$tempPortName' could not be removed: $($_.Exception.Message)"
                    }
                }
            }
        } else {
            Write-Verbose "[$Server] Temp port '$tempPortName' retained due to -NoCleanup."
        }

        # Verification
        $p = Get-Printer -Name $Name -ComputerName $Server -ErrorAction SilentlyContinue
        $success = $false
        $msg = ''
        if ($p -and $p.PortName -eq $finalPort) {
            $success = $true
            $msg = "Printer '$Name' now using port '$finalPort'."
        } else {
            $msg = "Verification failed: '$Name' is on port '$($p.PortName)' (expected '$finalPort')."
        }

        [pscustomobject]@{
            Server        = $Server
            PrinterName   = $Name
            FromPort      = $fromPort
            ToPort        = $finalPort
            TempPort      = $tempPortName
            PortRecreated = $recreated
            Success       = $success
            Message       = $msg
        }
    }
}

function Get-InstalledPrinters {
    <#
    .SYNOPSIS
    Lists per-user printer connections from one or more computers or PSSessions.

    .DESCRIPTION
    Enumerates registry keys under HKEY_USERS\<SID>\Printers\Connections on the target system(s) to report
    installed (connected) printers for each user profile. Works against:
    - One or more computer names via PowerShell remoting, or
    - One or more existing PSSessions.

    For each connection, returns a structured object with ComputerName, SID, (optionally) resolved User name,
    and the connection name (Printer).

    .PARAMETER ComputerName
    One or more remote computer names to query via PowerShell remoting. Accepts pipeline input.
    Requires WinRM connectivity and appropriate permissions.

    .PARAMETER Session
    One or more existing PSSessions to use for the query. Accepts pipeline input.

    .PARAMETER ResolveUser
    When specified, attempts to resolve each SID to an NTAccount (DOMAIN\User). If resolution fails,
    the SID is returned as-is.

    .PARAMETER ThrottleLimit
    Maximum number of concurrent remote calls when using -ComputerName. Default is 16.

    .PARAMETER Credential
    Specifies a user account that has permission to run the command on the remote computer(s).
    Only applicable with -ComputerName. When defined with the [Credential()] attribute and supplied
    without a value, you will be prompted for credentials.

    .INPUTS
    System.String, Microsoft.PowerShell.Commands.PSSession
    You can pipe computer names or PSSession objects.

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    Each object includes: ComputerName, SID, User (if -ResolveUser), and Printer.

    .EXAMPLES
    Example 1: Query a single computer
    PS> Get-InstalledPrinters -ComputerName 'PC01'

    Example 2: Query multiple computers with resolved users
    PS> 'PC01','PC02' | Get-InstalledPrinters -ResolveUser

    Example 3: Use existing sessions
    PS> $s = New-PSSession -ComputerName 'PC01','PC02'
    PS> Get-InstalledPrinters -Session $s

    Example 4: Capture and export
    PS> Get-InstalledPrinters -ComputerName 'PC01','PC02' -ResolveUser | Export-Csv printers.csv -NoTypeInformation

    Example 5: Use alternate credentials
    PS> Get-InstalledPrinters -ComputerName 'PC01','PC02' -Credential (Get-Credential)

    .NOTES
    Reads HKU (user hives) on target machines: HKEY_USERS\<SID>\Printers\Connections.
    Per-machine printers under HKLM are not included. Requires appropriate permissions.
    Errors during remote enumeration are reported via Write-Error but do not stop other targets.
    When -Session is used, -Credential is ignored because the PSSession encapsulates authentication.

    .LINK
    about_Remote
    .LINK
    about_Registry_Provider
    .LINK
    about_CommonParameters
    #>
    [CmdletBinding(DefaultParameterSetName = 'ComputerName')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ComputerName',
            ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('CN', 'Server')]
        [string[]]$ComputerName,

        [Parameter(Mandatory, ParameterSetName = 'Session',
            ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,

        [Parameter()]
        [switch]$ResolveUser,

        [Parameter(ParameterSetName = 'ComputerName')]
        [ValidateRange(1, 128)]
        [int]$ThrottleLimit = 16,

        [Parameter(ParameterSetName = 'ComputerName')]
        [Alias('Cred')]
        [pscredential]$Credential
    )

    begin {
        # Script block executed on remote targets
        $scriptBlock = {
            param([switch]$ResolveUser)

            try {
                $sids = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction Stop |
                    Where-Object { $_.PSChildName -match '^S-\d-\d+-(\d+-){1,14}\d+$' }

                foreach ($sid in $sids) {
                    $keyPath = "Registry::$($sid.Name)\Printers\Connections"
                    if (Test-Path $keyPath) {
                        foreach ($conn in (Get-ChildItem $keyPath -ErrorAction SilentlyContinue)) {
                            $user = $sid.PSChildName
                            if ($ResolveUser) {
                                try {
                                    $sidObj = [System.Security.Principal.SecurityIdentifier]$sid.PSChildName
                                    $user = $sidObj.Translate([System.Security.Principal.NTAccount]).Value
                                } catch { } # fall back to SID if translation fails
                            }

                            [pscustomobject]@{
                                PSTypeName   = 'AdventHealth.Prints.Connection'
                                ComputerName = $env:COMPUTERNAME
                                SID          = $sid.PSChildName
                                User         = $user
                                Printer      = $conn.PSChildName
                            }
                        }
                    }
                }
            } catch {
                Write-Error -ErrorRecord $_
            }
        }
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Session') {
            foreach ($s in $Session) {
                try {
                    Invoke-Command -Session $s -ScriptBlock $scriptBlock -ArgumentList $ResolveUser -ErrorAction Stop
                } catch {
                    Write-Error -ErrorRecord $_
                }
            }
        } else {
            try {
                Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $scriptBlock -ArgumentList $ResolveUser -ThrottleLimit $ThrottleLimit -ErrorAction Stop
            } catch {
                Write-Error -ErrorRecord $_
            }
        }
    }
}
function Get-ADPrinterGroups {
    <#
    .SYNOPSIS
    Gets the standard and default Active Directory groups for one or more printer names.

    .DESCRIPTION
    For each printer name, this function looks up two AD groups based on your naming convention:
    - Standard group: <Prefix><PrinterName>
    - Default group : <Prefix><PrinterName>-<DefaultMarker>

    By default, Prefix is 'ORL-PRN-' and DefaultMarker is 'DEF'. The function returns any groups
    that exist and ignores missing ones (reported via -Verbose). You can request additional AD
    properties with -Properties; the function will also include Description and info unless you
    request '*' (all properties).

    .PARAMETER PrinterName
    One or more printer names. Accepts pipeline input and property-based pipeline (Printer/Name).

    .PARAMETER Properties
    Additional AD group properties to retrieve. If '*' is specified, all properties are returned.
    Otherwise the function ensures 'Description' and 'info' are included.

    .PARAMETER Prefix
    The naming prefix for printer groups (default: 'ORL-PRN-').

    .PARAMETER DefaultMarker
    The suffix that identifies default groups (default: 'DEF').

    .PARAMETER Server
    Domain controller or AD LDS instance to query.

    .PARAMETER SearchBase
    Distinguished name (DN) that limits the search to a specific OU/container.

    .PARAMETER SearchScope
    Base, OneLevel, or Subtree.

    .INPUTS
    System.String

    .OUTPUTS
    Microsoft.ActiveDirectory.Management.ADGroup

    .EXAMPLES
    PS> Get-ADPrinterGroups -PrinterName 'IBJ9' -Verbose
    PS> 'IBJ9','HP-5100' | Get-ADPrinterGroups -Properties ManagedBy
    PS> Get-ADPrinterGroups -PrinterName 'X123' -Prefix 'MCO-PRN-' -DefaultMarker 'DEFAULT'
    PS> Get-ADPrinterGroups -PrinterName 'HP-Color' -Server 'dc01.contoso.com' -SearchBase 'OU=Print,DC=contoso,DC=com'

    .NOTES
    Requires RSAT ActiveDirectory module (Get-ADGroup).
    .LINK
    Get-ADGroup
    about_ActiveDirectory
    about_CommonParameters
    #>
    [CmdletBinding()]
    [OutputType([Microsoft.ActiveDirectory.Management.ADGroup])]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Printer', 'Name')]
        [ValidateNotNullOrEmpty()]
        [string[]]$PrinterName,

        [Parameter()]
        [string[]]$Properties = @(),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Prefix = 'ORL-PRN-',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$DefaultMarker = 'DEF',

        [Parameter()]
        [string]$Server,

        [Parameter()]
        [string]$SearchBase,

        [Parameter()]
        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [string]$SearchScope
    )

    begin {
        Import-Module -Name ActiveDirectory
        # normalize SearchScope for Get-ADGroup if provided
        $opt = @{}
        if ($PSBoundParameters.ContainsKey('Server')) { $opt.Server = $Server }
        if ($PSBoundParameters.ContainsKey('SearchBase')) { $opt.SearchBase = $SearchBase }
        if ($PSBoundParameters.ContainsKey('SearchScope')) { $opt.SearchScope = $SearchScope }
        
        # ---- Normalize $Properties ----
        if ($Properties -is [string]) {
            $Properties = $Properties -split '\s*,\s*' | Where-Object { $_ }
        }
        $Properties = @($Properties)  # ensure array
    
        $effectiveProps = if ($Properties -contains '*') { '*' } else {
            @($Properties + 'Description' + 'info' | Select-Object -Unique)
        }
    }

    process {
        foreach ($name in $PrinterName) {
            if (-not $name) { continue }

            # Build the two identities we want to try
            $ids = @(
                ('{0}{1}' -f $Prefix, $name),
                ('{0}{1}-{2}' -f $Prefix, $name, $DefaultMarker)
            )

            $output = foreach ($id in $ids) {
                try {
                    Write-Verbose "Querying group: $id"
                    Get-ADGroup -Identity $id -Properties $effectiveProps @opt -ErrorAction Stop
                } catch {
                    Write-Verbose "Group not found or inaccessible: $id ($($_.Exception.Message))"
                }
            }
        }
        if (-not $output) {
            $filterString = "*-PRN-$name*"
            $output = Get-ADGroup -Filter { Name -like $filterString } -Properties $effectiveProps @opt
        } 
        return $output
    }
}


function Get-ADPrinters {
    <#
    .SYNOPSIS
    Retrieves Active Directory printer objects from specified print servers and exports the results to Excel.

    .DESCRIPTION
    The Get-ADPrinters function queries a list of print servers for installed printers, then searches Active Directory for matching printQueue objects. It supports filtering by creation date, either by specifying a number of months ago or a specific date. The results are exported to an Excel file, with a prompt to close Excel if it is running.

    .PARAMETER Domain
    Specifies the Active Directory domain to query. Defaults to 'flhosp.net'. If set to 'multihosp.net', the function uses a different set of print servers.

    .PARAMETER CreatedMonthsAgo
    Filters printers created within the specified number of months from the current date. Mutually exclusive with CreatedAfter.

    .PARAMETER CreatedAfter
    Filters printers created after the specified date (in a format recognized by Get-Date). Mutually exclusive with CreatedMonthsAgo.

    .PARAMETER Servers
    An array of print server names to query. Defaults to a predefined list based on the selected domain.

    .EXAMPLE
    Get-ADPrinters -CreatedMonthsAgo 6

    Retrieves printers created in the last 6 months from the default domain and exports the results to Excel.

    .EXAMPLE
    Get-ADPrinters -CreatedAfter '2024-01-01' -Domain 'multihosp.net'

    Retrieves printers created after January 1, 2024, from the 'multihosp.net' domain and exports the results to Excel.

    .NOTES
    - Requires the ActiveDirectory and ImportExcel modules.
    - Prompts the user to close Excel before exporting data.
    - Output is saved to "$env:HOMEPATH\Documents\AD_Printer_Export.xlsx".
    #>


    [CmdletBinding(DefaultParameterSetName = 'CreatedAfter')]
    param(
        [string]$Domain = 'flhosp.net',
        [Parameter(ParameterSetName = 'CreatedMonthsAgo')]
        [Int32]$CreatedMonthsAgo,
        [Parameter(ParameterSetName = 'CreatedAfter')]
        [string]$CreatedAfter,
        [string[]]$Servers = ('FHOSVMWPRN001', `
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
    )

    if ($Domain -eq 'multihosp.net') {
        $Servers = ('ADCVPRNMHDMS001', `
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
    }

    # Initialize variables.
    $results = @()
    $printers = $Servers | ForEach-Object -Parallel { Get-Printer -ComputerName $_ | Sort-Object | Select-Object Name }
    $props = ('whenCreated', 'ServerName', 'Description', 'Location', 'PortName')
    
    if ($CreatedMonthsAgo) {
        $WorkSheetName = 'Last {0} Months' -f $CreatedMonthsAgo
        foreach ($printer in $printers) {
            $filter = "ObjectClass -like 'printQueue' -and Name -like '*$($printer.Name)*'"
            $printerObj = Get-ADObject -Filter $filter -Server $Domain -Properties $props | Where-Object whenCreated -GT (Get-Date).AddMonths("-$CreatedMonthsAgo").Date
            if ($printerObj.whenCreated -gt (Get-Date).AddMonths("-$CreatedMonthsAgo").Date) {
                $results += [PSCustomObject]@{
                    'Name'        = $printerObj.Name
                    'Port'        = $printerObj.PortName[0]
                    'Server'      = $printerObj.ServerName
                    'Create Date' = $printerObj.whenCreated
                    'Location'    = $printerObj.Location
                    'Description' = $printerObj.Description
                }
            }
        }

        Stop-ExcelProcess
        $results | Export-Excel -Path "$env:HOMEPATH\Documents\AD_Printer_Export.xlsx" -WorksheetName $WorkSheetName -AutoSize -TableName $WorkSheetName -ClearSheet
        Write-Host "Output saved to $env:HOMEPATH\Documents\AD_Printer_Export.xlsx"
    } elseif ($CreatedAfter) {
        $WorkSheetName = 'Created After {0}' -f $CreatedAfter
        foreach ($printer in $printers) {
            $filter = "ObjectClass -like 'printQueue' -and Name -like '*{0}*'" -f $printer.Name
            $printerObj = Get-ADObject -Filter $filter -Server $Domain -Properties $props | Where-Object whenCreated -GT (Get-Date $CreatedAfter).Date
            if ($printerObj.whenCreated -gt (Get-Date $CreatedAfter).Date) {
                $results += [PSCustomObject]@{
                    'Name'        = $printerObj.Name
                    'Port'        = $printerObj.PortName[0]
                    'Server'      = $printerObj.ServerName
                    'Create Date' = $printerObj.whenCreated
                    'Location'    = $printerObj.Location
                    'Description' = $printerObj.Description
                }
            }
        }

        Stop-ExcelProcess
        $results | Export-Excel -Path "$env:HOMEPATH\Documents\AD_Printer_Export.xlsx" -WorksheetName $WorkSheetName -AutoSize -TableName $WorkSheetName -ClearSheet
        Write-Host "Output saved to $env:HOMEPATH\Documents\AD_Printer_Export.xlsx"
    } else {
        foreach ($printer in $printers) {
            $filter = "ObjectClass -like 'printQueue' -and Name -like '*" + $printer.Name + "*'"
            $printerObj = Get-ADObject -Filter $filter -Server $Domain -Properties $props
            if ($printerObj) {
                $results += [PSCustomObject]@{
                    'Name'        = $printerObj.Name
                    'Port'        = $printerObj.PortName[0]
                    'Server'      = $printerObj.ServerName
                    'Create Date' = $printerObj.whenCreated
                    'Location'    = $printerObj.Location
                    'Description' = $printerObj.Description
                }
            }
        }
        Stop-ExcelProcess
        $results | Export-Excel -Path "$env:HOMEPATH\Documents\AD_Printer_Export.xlsx" -WorksheetName 'All Printers' -AutoSize -TableName 'All Printers' -ClearSheet
        Write-Host "Output saved to $env:HOMEPATH\Documents\AD_Printer_Export.xlsx"
    }
}

function Repair-Printer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $Name,

        [Parameter()]
        [string[]]
        $Servers = @('ADCVPRNMHDMS001', `
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
        $printerHash = @{}
        $props = @('ComputerName', 'Name', 'Location', 'Comment', 'DriverName', 'ShareName', 'Shared', 'PortName')
    }

    process {
        $printer = $Servers | ForEach-Object -Parallel { 
            Get-Printer -ComputerName $_ -Name $Name -ErrorAction SilentlyContinue | Select-Object $props 
        }

        if ($printer) {
            $printer.psobject.Properties | ForEach-Object {
                $printerHash[$_.Name] = $_.Value
            }
            try {
                # Remove-Printer isn't working for some reason...
                $removeResult = Remove-Printer -ComputerName $printer.ComputerName -Name $printer.Name -ErrorAction Stop
            } catch {
                Write-Error "Failed to remove $($printer.Name) from $($printer.ComputerName): ($($_.Exception.Message))"
            }

            if ($removeResult) {
                Add-Printer @printerHash
            }
        }
    }
}

function Set-InitialPrinterConfig {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory)]
        [string]$Server,

        [Parameter(Position = 1, Mandatory)]
        [string]$PrinterName,

        [Parameter(Position = 2)]
        [string]
        $PrinterModel
    )

    begin {
        $printProperties = @{
            'Config:DuplexUnit'         = 'Installed'
            'Config:JobStorageControl'  = 'auto'
            'Config:AccessoryOutputBin' = 'NoOutputBin'
            'Config:SecurePrintControl' = 'auto'
            'Config:TintTestingControl' = 'disable'
        }

        function Get-AvailablePrintDrivers {
            param ($Server, $PrinterModel)

            $installedDrivers = Get-PrinterDriver -ComputerName $Server | Where-Object { $_.Name -like "*$PrinterModel*" }
            
            if ($installedDrivers.Count -eq 1) {
                $selectedDriver = $installedDrivers.Name
            }

            if ($installedDrivers.Count -gt 1) {
                Write-Host "Multiple drivers available for $PrinterModel on $Server"
                $idx = 0
                foreach ($drvr in $installedDrivers) {
                    Write-Host "$($idx): $($drvr.Name)"
                    $idx++
                }
                [int]$selectedIdx = Read-Host 'Please select the preferred driver'
                $selectedDriver = $installedDrivers[$selectedIdx].Name
            }

            if ($installedDrivers.Count -eq 0) {
                throw "No driver found for $PrinterModel on $Server."
            }

            return $selectedDriver
        }

        Write-Verbose "Configuring printer '$PrinterName' on server '$Server'..."
    }

    process {
        $results = @()

        if ($PrinterModel) {
            try {
                $selectedPrintDriver = Get-AvailablePrintDrivers $Server $PrinterModel -ErrorAction Stop
                try {
                    $currentPrintDriver = Get-Printer -ComputerName $Server -Name $PrinterName -ErrorAction Stop | Select-Object -ExpandProperty DriverName 
                } catch {
                    Write-Error "[Get-Printer]: $($_.Exception.Message)"
                }

                if ($selectedPrintDriver -ne $currentPrintDriver) {
                    Set-Printer -ComputerName $Server -Name $PrinterName -DriverName $selectedPrintDriver -ErrorAction Stop
                    Write-Verbose "✔ Set $PrinterName to use $selectedPrintDriver."
                } else { Write-Verbose "✔ $PrinterName is already set to use $currentPrintDriver." }

                $results += [PSCustomObject]@{
                    Setting = 'Driver'
                    Value   = $selectedPrintDriver
                    Status  = 'Success'
                }
            } catch {
                Write-Error "[Set-Printer]: $($_.Exception.Message)" 
                return
            }
        }

        try {
            Set-PrintConfiguration `
                -ComputerName $Server `
                -PrinterName $PrinterName `
                -DuplexingMode OneSided `
                -PaperSize Letter `
                -Collate $true `
                -Color $true

            Write-Verbose '✔ Base print configuration applied successfully.'

            $results += [PSCustomObject]@{
                Setting = 'BaseConfiguration'
                Value   = 'Applied'
                Status  = 'Success'
            }
        } catch {
            Write-Error "Failed to apply base print configuration: $_"
        }

        foreach ($property in $printProperties.GetEnumerator()) {
            try {
                Set-PrinterProperty `
                    -ComputerName $Server `
                    -PrinterName $PrinterName `
                    -PropertyName $property.Name `
                    -Value $property.Value `
                    -ErrorAction Stop

                Write-Verbose "✔ Set $($property.Name) = $($property.Value)"

                $results += [PSCustomObject]@{
                    Setting = $property.Name
                    Value   = $property.Value
                    Status  = 'Success'
                }
            } catch {
                Write-Verbose "✖ Failed to set $($property.Name): $_"

                $results += [PSCustomObject]@{
                    Setting = $property.Name
                    Value   = $property.Value
                    Status  = 'Failed'
                }
            }
        }
    }

    end {
        Write-Verbose 'Printer configuration complete.'
        return $results
    }
}