if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
    Write-Host 'ImportExcel module not found. Installing...'
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    Write-Host 'ImportExcel module installed.'
} else { Import-Module ImportExcel }

# Use Type Accelerator to allow PSSession type.
[PowerShell].Assembly.GetType('System.Management.Automation.TypeAccelerators')::add('PSSession', 'System.Management.Automation.Runspaces.PSSession')

function Stop-ExcelProcess {
    <#
    .SYNOPSIS
    Stops any running Microsoft Excel processes, with optional confirmation and WhatIf support.

    .DESCRIPTION
    Stop-ExcelProcess detects running instances of Microsoft Excel and stops them. By default, if Excel is running,
    the function prompts for confirmation using ShouldContinue. You can bypass the prompt with -Force, or rely on
    PowerShell’s standard -WhatIf / -Confirm behavior via SupportsShouldProcess.

    The function returns a Boolean indicating the outcome:
    - $true  : Excel was not running or was successfully stopped.
    - $false : The user cancelled or an error occurred while stopping Excel.

    .PARAMETER Force
    Skips the interactive confirmation prompt and attempts to stop Excel immediately. Still respects -WhatIf/-Confirm.

    .INPUTS
    None. You cannot pipe input to this function.

    .OUTPUTS
    System.Boolean
    Returns $true on success (or when Excel is already closed), $false on cancellation or error.

    .EXAMPLE
    PS> Stop-ExcelProcess
    Prompts to close running Excel processes, then stops them if confirmed.

    .EXAMPLE
    PS> Stop-ExcelProcess -Force
    Immediately stops any running Excel processes without prompting.

    .EXAMPLE
    PS> Stop-ExcelProcess -WhatIf
    Shows what would happen if the function ran, without making changes.

    .NOTES
    Requires Windows when targeting Microsoft Excel as a desktop application.
    Uses Get-Process 'EXCEL' and Stop-Process -Force.
    Integrates with -WhatIf and -Confirm via SupportsShouldProcess.

    .LINK
    about_Comment_Based_Help
    .LINK
    about_CommonParameters
    #>

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
            $output = Get-ADGroup -Filter { Name -like $filterString }
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

function Get-Locked {
    <#
    .SYNOPSIS
    Checks whether one or more remote computers are currently locked; optionally monitors until they become locked.

    .DESCRIPTION
    Determines lock state by testing for the presence of the LogonUI process on the target computer(s).
    Returns structured objects with ComputerName, IsLocked, and Timestamp. When -Monitor is specified,
    the function repeatedly checks until the computer reports as locked or an optional timeout elapses.

    Supports querying by -ComputerName (with optional -Credential) or by existing -Session objects.

    .PARAMETER ComputerName
    One or more remote computers to query via PowerShell remoting. Accepts pipeline input.

    .PARAMETER Session
    One or more existing PSSessions to query.

    .PARAMETER Credential
    Credentials used when connecting to remote computers with -ComputerName. Ignored when -Session is used.

    .PARAMETER Monitor
    When specified, repeatedly checks each target until it becomes locked (LogonUI is present),
    or until -TimeoutSeconds (if provided) elapses.

    .PARAMETER IntervalSeconds
    Polling interval for -Monitor. Default: 5 seconds.

    .PARAMETER TimeoutSeconds
    Maximum time to wait (in seconds) for -Monitor. Default: 0 (no timeout).

    .PARAMETER ThrottleLimit
    Maximum concurrent remote calls for -ComputerName. Default: 16.

    .INPUTS
    System.String, Microsoft.PowerShell.Commands.PSSession
    You can pipe computer names or PSSession objects.

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    Properties: ComputerName (String), IsLocked (Boolean), Timestamp (DateTime)
    #>
    [CmdletBinding(DefaultParameterSetName = 'ComputerName')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'ComputerName',
            ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory, ParameterSetName = 'Session',
            ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,

        [Parameter(ParameterSetName = 'ComputerName')]
        [pscredential]$Credential,

        [Parameter()]
        [switch]$Monitor,

        [Parameter()]
        [ValidateRange(1, 3600)]
        [int]$IntervalSeconds = 5,

        [Parameter()]
        [ValidateRange(0, 86400)]
        [int]$TimeoutSeconds = 0,

        [Parameter(ParameterSetName = 'ComputerName')]
        [ValidateRange(1, 128)]
        [int]$ThrottleLimit = 16
    )

    begin {
        # Remote test: returns $true if LogonUI is running (locked), else $false
        $lockTestScript = {
            $null -ne (Get-Process -Name 'LogonUI' -ErrorAction SilentlyContinue)
        }
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Session') {
            foreach ($s in $Session) {
                if ($Monitor) {
                    $deadline = if ($TimeoutSeconds -gt 0) { (Get-Date).AddSeconds($TimeoutSeconds) } else { $null }
                    while ($true) {
                        try {
                            $isLocked = Invoke-Command -Session $s -ScriptBlock $lockTestScript -ErrorAction Stop
                        } catch {
                            Write-Error -Message "Failed to query $($s.ComputerName): $($_.Exception.Message)" -ErrorRecord $_
                            break
                        }
                        if ($isLocked) {
                            [pscustomobject]@{
                                ComputerName = $s.ComputerName
                                IsLocked     = $true
                                Timestamp    = Get-Date
                            }
                            break
                        }
                        if ($deadline -and (Get-Date) -ge $deadline) {
                            [pscustomobject]@{
                                ComputerName = $s.ComputerName
                                IsLocked     = $false
                                Timestamp    = Get-Date
                            }
                            break
                        }
                        Start-Sleep -Seconds $IntervalSeconds
                    }
                } else {
                    try {
                        $isLocked = Invoke-Command -Session $s -ScriptBlock $lockTestScript -ErrorAction Stop
                        [pscustomobject]@{
                            ComputerName = $s.ComputerName
                            IsLocked     = [bool]$isLocked
                            Timestamp    = Get-Date
                        }
                    } catch {
                        Write-Error -Message "Failed to query $($s.ComputerName): $($_.Exception.Message)" -ErrorRecord $_
                    }
                }
            }
        } else {
            if ($Monitor) {
                foreach ($cn in $ComputerName) {
                    $deadline = if ($TimeoutSeconds -gt 0) { (Get-Date).AddSeconds($TimeoutSeconds) } else { $null }
                    while ($true) {
                        try {
                            $isLocked = Invoke-Command -ComputerName $cn -Credential $Credential `
                                -ScriptBlock $lockTestScript -ErrorAction Stop
                        } catch {
                            Write-Error -Message "Failed to query $($cn): $($_.Exception.Message)" -ErrorRecord $_
                            break
                        }
                        if ($isLocked) {
                            [pscustomobject]@{
                                ComputerName = $cn
                                IsLocked     = $true
                                Timestamp    = Get-Date
                            }
                            break
                        }
                        if ($deadline -and (Get-Date) -ge $deadline) {
                            [pscustomobject]@{
                                ComputerName = $cn
                                IsLocked     = $false
                                Timestamp    = Get-Date
                            }
                            break
                        }
                        Start-Sleep -Seconds $IntervalSeconds
                    }
                }
            } else {
                try {
                    Invoke-Command -ComputerName $ComputerName -Credential $Credential `
                        -ThrottleLimit $ThrottleLimit -ScriptBlock $lockTestScript -ErrorAction Stop |
                        ForEach-Object {
                            [pscustomobject]@{
                                ComputerName = $_.PSComputerName   # ✅ FIX: grab property, not variable
                                IsLocked     = [bool]$_
                                Timestamp    = Get-Date
                            }
                        }
                } catch {
                    Write-Error -Message "Failed one or more queries: $($_.Exception.Message)" -ErrorRecord $_
                }
            }
        }
    }
}

function Enter-PSSessionWithCred {
    <#
    .SYNOPSIS
    Initiates a remote PowerShell session to a specified computer using a global credential.

    .DESCRIPTION
    The Enter-PSSessionWithCred function simplifies the process of starting a remote PowerShell session by using a predefined global credential stored in `$global:cred`. This is useful for administrators who frequently connect to remote systems and want to avoid repeated credential prompts.

    .PARAMETER ComputerName
    Specifies the name of the remote computer to connect to. This parameter is mandatory.

    .EXAMPLE
    Enter-PSSessionWithCred -ComputerName 'WS-12345'

    Starts a remote PowerShell session with the computer named 'WS-12345' using the global credential.

    .NOTES
    - Requires PowerShell remoting to be enabled on the target computer.
    - Assumes `$global:cred` is already defined and contains valid credentials.
    #>

    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    Enter-PSSession -ComputerName $ComputerName -Credential $global:cred
}

function Get-InstalledSoftware {
    <#
    .SYNOPSIS
    Retrieves a list of installed software from remote computers using the registry.

    .DESCRIPTION
    This function safely queries both 64-bit and 32-bit uninstall registry locations to retrieve software data. If the software was installed via MSI, its GUID is returned.

    .PARAMETER ComputerName
    One or more remote computers to query.

    .PARAMETER Session
    One or more PowerShell sessions to use instead of computer names.

    .PARAMETER Filter
    Optional string to filter software by partial name match.

    .EXAMPLE
    Get-InstalledSoftware -ComputerName 'WS-12345' -Filter 'Microsoft'
    #>

    [CmdletBinding(DefaultParameterSetName = 'ComputerName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ComputerName')]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true, ParameterSetName = 'Session')]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,

        [string]$Filter
    )

    $scriptBlock = {
        param($filter)

        function Get-InstalledAppsFromRegistry {
            param ($regPath)

            if (Test-Path $regPath) {
                Get-ChildItem -Path $regPath | ForEach-Object {
                    $key = $_
                    $props = Get-ItemProperty -Path $key.PSPath
                    if ($props.DisplayName) {
                        $subkey = $key.PSChildName
                        $isGuid = $subkey -match '^\{[0-9A-Fa-f\-]{36}\}$'

                        [PSCustomObject]@{
                            Name            = $props.DisplayName
                            Version         = $props.DisplayVersion
                            Publisher       = $props.Publisher
                            InstallDate     = $props.InstallDate
                            UninstallString = $props.UninstallString
                            GUID            = if ($isGuid) { $subkey } else { $null }
                            RegistryPath    = $key.PSPath
                        }
                    }
                }
            }
        }

        $paths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        )

        $apps = foreach ($path in $paths) {
            Get-InstalledAppsFromRegistry -regPath $path
        }

        if ($filter) {
            $apps | Where-Object { $_.Name -like "*$filter*" } | Sort-Object Name
        } else {
            $apps | Sort-Object Name
        }
    }

    switch ($PSCmdlet.ParameterSetName) {
        'ComputerName' {
            Invoke-Command -ComputerName $ComputerName -Credential $global:cred -ScriptBlock $scriptBlock -ArgumentList $Filter
        }
        'Session' {
            Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $Filter
        }
    }
}




function Remove-Software {
    <#
    .SYNOPSIS
    Uninstalls software on a remote computer using a GUID or uninstall string.

    .DESCRIPTION
    Performs a silent uninstall using either the MSI product code (GUID) or a full uninstall command line string. Uses PowerShell remoting and includes logging and error handling.

    .PARAMETER ComputerName
    One or more target computers to uninstall from.

    .PARAMETER Session
    One or more existing PowerShell remoting sessions.

    .PARAMETER GUID
    The product GUID (MSI ProductCode) of the software to uninstall.

    .PARAMETER UninstallString
    A complete uninstall command string (non-MSI uninstallers).

    .PARAMETER Credential
    Optional credential for Invoke-Command. Defaults to $global:cred.

    .EXAMPLE
    Remove-Software -ComputerName 'WS-12345' -GUID '{12345678-1234-1234-1234-1234567890AB}'

    .EXAMPLE
    Remove-Software -ComputerName 'WS-12345' -UninstallString '"C:\Program Files\App\uninstall.exe" /quiet'
    #>

    [CmdletBinding(DefaultParameterSetName = 'ComputerName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ComputerName')]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $true, ParameterSetName = 'Session')]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,

        [Parameter()]
        [string]$GUID,

        [Parameter()]
        [string]$UninstallString,

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential = $global:cred
    )

    if (-not $GUID -and -not $UninstallString) {
        throw 'You must specify either -GUID or -UninstallString.'
    }

    if ($GUID -and $GUID -notmatch '^\{[0-9A-Fa-f\-]{36}\}$') {
        throw 'Invalid GUID format. Expected: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}'
    }

    $scriptBlock = {
        param($guid, $uninstallString)

        $logPath = "C:\Windows\Temp\Uninstall_$($env:COMPUTERNAME)_$((Get-Date).ToString('yyyyMMddHHmmss')).log"
        Start-Transcript -Path $logPath -Append | Out-Null

        try {
            if ($guid) {
                Start-Process -FilePath 'msiexec.exe' -ArgumentList '/x', $guid, '/qn', 'REBOOT=ReallySuppress' -Wait -NoNewWindow -ErrorAction Stop
            } elseif ($uninstallString) {
                Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `"$uninstallString`"" -Wait -NoNewWindow -ErrorAction Stop
            }

            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Identifier   = if ($guid) { $guid } else { $uninstallString }
                Status       = 'Success'
                LogPath      = $logPath
            }
        } catch {
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Identifier   = if ($guid) { $guid } else { $uninstallString }
                Status       = 'Failed'
                Error        = $_.Exception.Message
                LogPath      = $logPath
            }
        } finally {
            Stop-Transcript | Out-Null
        }
    }

    switch ($PSCmdlet.ParameterSetName) {
        'ComputerName' {
            Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $scriptBlock -ArgumentList $GUID, $UninstallString
        }
        'Session' {
            Invoke-Command -Session $Session -ScriptBlock $scriptBlock -ArgumentList $GUID, $UninstallString
        }
    }
}

function Get-EODTime {
    <#
    .SYNOPSIS
    Calculates the projected end-of-day time based on hours worked and a given start time.

    .DESCRIPTION
    The Get-EODTime function determines the estimated end-of-day time for a work shift based on the number of hours already worked and a specified start time. If the hours worked are 40 or less, it calculates the time needed to reach 40 hours. If more than 40 hours have been worked, it calculates the time needed to reach 80 hours. A 30-minute buffer is added to the result.

    .PARAMETER HoursWorked
    The number of hours already worked. This value is used to calculate the remaining time needed to reach 40 or 80 hours.

    .PARAMETER StartTime
    The starting time (in a format accepted by Get-Date) from which the remaining hours will be added to compute the end-of-day time.

    .EXAMPLE
    Get-EODTime -HoursWorked 32 -StartTime '2025-07-28 08:00'

    Returns the time when the user will reach 40 hours, starting from 8:00 AM, with a 30-minute buffer.

    .EXAMPLE
    Get-EODTime -HoursWorked 45 -StartTime '2025-07-28 09:00'

    Returns the time when the user will reach 80 hours, starting from 9:00 AM, with a 30-minute buffer.

    .NOTES
    - Useful for shift planning and time tracking.
    - Adds a 0.5-hour (30-minute) buffer to the calculated end time.
    #>

    param(
        [Parameter(Mandatory = $true)]
        [string]$HoursWorked,
        [Parameter(Mandatory = $true)]
        [string]$StartTime
    )
    if ($HoursWorked -le 40) {
        Write-Output (Get-Date $StartTime).AddHours(40 - $HoursWorked + 0.5)
    } else {
        Write-Output (Get-Date $StartTime).AddHours(80 - $HoursWorked + 0.5)
    }
}

function Get-ADGroupMemberships {
    <#
    .SYNOPSIS
    Retrieves the Active Directory groups that a specified user or computer is a member of.

    .DESCRIPTION
    The Get-ADGroupMemberships function queries Active Directory to return the group memberships for a given user or computer identity. 
    It attempts to locate the object first as a user, and if that fails, as a computer. 
    Optionally, specific group properties can be returned.

    .PARAMETER Identity
    Specifies the Active Directory user or computer account to retrieve group memberships for.
    This parameter is required.

    .PARAMETER Properties
    Specifies the properties of each group to return.
    If omitted, all default properties of the group object are returned.

    .EXAMPLE
    Get-ADGroupMemberships -Identity jdoe

    Retrieves all groups that the user account "jdoe" is a member of.

    .EXAMPLE
    Get-ADGroupMemberships -Identity srv-web01$ -Properties Name,Description

    Retrieves the Name and Description of all groups that the computer account "srv-web01$" is a member of.

    .NOTES
    Author: Scott C. Court
    Requires: ActiveDirectory module

    #>

    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Identity,
        [string[]]$Properties = @()
    )

    # If user passed a single comma-separated string, split it
    if ($Properties -is [string]) {
        $Properties = $Properties -split '\s*,\s*' | Where-Object { $_ }
    }
    $Properties = @($Properties)             # ensure array
    if ($Properties -contains '*') {
        $effectiveProps = '*'
    } else {
        $effectiveProps = @($Properties + 'MemberOf' | Select-Object -Unique)
    }

    $splat = @{
        Identity   = $Identity
        Properties = $effectiveProps
    }

    try {
        $adObj = Get-ADComputer @splat -ErrorAction Ignore | Select-Object -ExpandProperty MemberOf 
    } catch {}
    try {
        if (-not $adObj) {
            $adObj = Get-ADUser @splat -ErrorAction Stop | Select-Object -ExpandProperty MemberOf
        }
    } catch {
        Write-Error "AD Object not found for $($splat.Identity). ($($_.Exception.Message))"
    }
    
    Write-Host "$Identity Group Memberships:"
    if ($Properties) {
        $groups = $adObj | ForEach-Object {
            Get-ADGroup $_ -Properties $Properties | Select-Object -Property $Properties
        } | Sort-Object $Properties
        return $groups
    } else {
        $groups = $adObj | ForEach-Object {
            Get-ADGroup $_
        } | Sort-Object Name
        return $groups
    }
}

function Get-ADGroupMembers {
    <#
    .SYNOPSIS
    Retrieves members of Active Directory groups using group names, a list file, or a directory structure, and optionally exports the results to Excel.

    .DESCRIPTION
    The Get-ADGroupMembers function collects members of AD groups based on one of three input methods:
    - A list of group names.
    - A file path containing group names.
    - A directory structure where folder names are used to infer group names.

    It supports optional exclusion filters and exports the results to Excel using the ImportExcel module. The function is designed to work with the 'multihosp.net' domain and handles both user and computer objects as group members.

    .PARAMETER GroupNames
    An array of group name patterns to search for in Active Directory. Used with the 'GroupNames' parameter set.

    .PARAMETER GroupListPath
    Path to a file containing group names, one per line. Used with the 'GroupListPath' parameter set.

    .PARAMETER DirectoryPath
    Path to a directory structure where subfolder names are used to construct group name patterns. Used with the 'DirectoryPath' parameter set.

    .PARAMETER Exclude
    An array of string patterns to exclude from the group name search results.

    .EXAMPLE
    Get-ADGroupMembers -GroupNames 'IT-Admin', 'HR-Users'

    Retrieves members of groups matching 'IT-Admin' and 'HR-Users' and displays them in the console.

    .EXAMPLE
    Get-ADGroupMembers -GroupListPath 'C:\Groups\groupnames.txt' -Exclude 'Test', 'Temp'

    Retrieves members of groups listed in the file, excluding any that match 'Test' or 'Temp', and displays them in the console.

    .EXAMPLE
    Get-ADGroupMembers -DirectoryPath 'C:\DepartmentFolders'

    Infers group names from subfolder names and exports the results to Excel files named after each parent folder.

    .NOTES
    - Requires the ActiveDirectory and ImportExcel modules.
    - Uses the 'multihosp.net' domain for all AD queries.
    - Group members are resolved as either user or computer objects.
    - Excel export is performed when using DirectoryPath or GroupListPath.
    - Output is saved to: $env:HOMEPATH\Documents\AD Groups
    #>

    [CmdletBinding(DefaultParameterSetName = 'GroupNames')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'GroupNames', Position = 0)]
        [string[]]$GroupNames,
        [Parameter(Mandatory = $true, ParameterSetName = 'GroupListPath')]
        [string]$GroupListPath,
        [Parameter(Mandatory = $true, ParameterSetName = 'DirectoryPath')]
        [string]$DirectoryPath,
        [string[]]$Exclude
    )

    $outputPath = "$env:HOMEPATH\Documents\AD Groups"
    $domainName = 'multihosp.net'
    $filterArray = @()
    $results = @()

    if ($Exclude) {
        $excludeArray = @()
        $Exclude | ForEach-Object -Parallel { $excludeArray += "(Name -notlike '*$_*')" }
        $excludeString = $excludeArray -join ' -and '
    }

    if ($GroupNames) {
        $GroupNames | ForEach-Object { $filterArray += "(Name -like '*$_*')" }
    } elseif ($GroupListPath) {
        $content = Get-Content -Path $GroupListPath
        $content | ForEach-Object { $filterArray += "(Name -like '*$_*')" }
    } elseif ($DirectoryPath) {
        $parentPaths = Get-ChildItem -Path $DirectoryPath -Attributes Directory
        foreach ($parentPath in $parentPaths) {
            $filterArray = @()
            $results = @()
            $parentPathName = $parentPath.Name
            $folderNames = Get-ChildItem -Path $parentPath -Attributes Directory | ForEach-Object -Parallel { (Get-Item -Path $_).Parent.Name + '-' + (Get-Item -Path $_).Name }
            $folderNames | ForEach-Object -Parallel { if (Get-ADGroup -Filter "(Name -like 'Prm-Fil-$_')" -Server $domainName -Properties Members -ErrorAction SilentlyContinue) { $filterArray += "(Name -like 'Prm-Fil-$_')" } }

            if ($filterArray.Count -gt 0) {
                foreach ($filter in $filterArray) {
                    if ($Exclude) {
                        $filter += " -and ($excludeString)"
                    }
                    $groups = Get-ADGroup -Filter $filter -Server $domainName -Properties Members -ErrorAction SilentlyContinue
                    $results += foreach ($group in $groups) {
                        Write-Host 'Collecting members of' $group.Name
                        foreach ($memberDN in $group.Members) {
                            try {
                                $user = Get-ADUser -Identity $memberDN -Server $domainName -Properties DisplayName -ErrorAction SilentlyContinue
                                if ($user) {
                                    [PSCustomObject]@{
                                        'Group Name'          = $group.Name
                                        'Member Display Name' = $user.DisplayName
                                    }
                                }
                            } catch { Write-Host "$memberDN is not a User object." }
                        }
                    }
                }
                # Export to Excel (requires ImportExcel module)
                $results | Export-Excel -Path "$outputPath\$parentPathName.xlsx" -WorksheetName 'Group Members' -AutoSize -TableName 'GroupMembers' -ClearSheet
                Write-Host "Export complete: $outputPath\$parentPathName.xlsx"
                return $results
            }
        }
        # Export to Excel (requires ImportExcel module)
        $outputPath += '\AD_Groups_Report.xlsx'
        $results | Export-Excel -Path $outputPath -WorksheetName 'Group Members' -AutoSize -TableName 'GroupMembers' -ClearSheet

        Write-Host "Export complete: $outputPath"

    }

    if (-not $DirectoryPath) {
        foreach ($filter in $filterArray) {
            if ($Exclude) {
                $filter += " -and ($excludeString)"
            }
            $groups = Get-ADGroup -Filter $filter -Server $domainName -Properties Members -ErrorAction SilentlyContinue
            $results += foreach ($group in $groups) {
                Write-Host 'Collecting members of' $group.Name
                foreach ($memberDN in $group.Members) {
                    try { $user = Get-ADUser -Identity $memberDN -Server $domainName -Properties DisplayName -ErrorAction SilentlyContinue } catch {}
                    if (-not $user) { $computer = Get-ADComputer -Identity $memberDN -Server $domainName -Properties DisplayName -ErrorAction SilentlyContinue }
                    if ($user) {
                        [PSCustomObject]@{
                            'Group Name'          = $group.Name
                            'Member Display Name' = $user.DisplayName
                        }
                    } elseif ($computer) {
                        [PSCustomObject]@{
                            'Group Name'          = $group.Name
                            'Member Display Name' = $computer.DisplayName
                        }
                    }
                }
            }
        }
        return $results | Format-List
        
    }
}

function Get-OfflineDevices {
    <#
    .SYNOPSIS
    Checks the online status of devices listed in an Excel file and exports the results to a new Excel file.

    .DESCRIPTION
    The Get-OfflineDevices function reads a list of devices from an Excel file, extracts and cleans the IP address or port information, and then pings each device to determine if it is online or offline. The results are split into two categories—online and offline—and exported to separate worksheets in a single Excel file. If Excel is running, the user is prompted to close it before export.

    .PARAMETER Path
    Specifies the path to the Excel file containing device information. The file must include a column named 'Port' with IP addresses or hostnames.

    .EXAMPLE
    Get-OfflineDevices -Path 'C:\Network\devices.xlsx'

    Checks the connectivity of devices listed in the specified Excel file and exports the results to `Offline_Devices.xlsx` in the user's Documents folder.

    .NOTES
    - Requires the ImportExcel module.
    - Uses the `Test-Connection` cmdlet to determine device availability.
    - Prompts the user to close Excel before writing the output.
    - Output is saved to: `$env:HOMEPATH\Documents\Offline_Devices.xlsx`
    - Calls `Stop-ExcelProcess` to ensure Excel is not running before export.
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Path
    )

    $content = Import-Excel -Path $Path
    $onlineDevices = @()
    $offlineDevices = @()

    $content | ForEach-Object -Parallel {
        $cleanPort = $_.Port -replace '_\d+$', ''
        if ($cleanPort -match '^(?:\d{1,3}\.){3}\d{1,3}$') {
            if ( Test-Connection -ComputerName $cleanPort -Count 2 ) {
                # Write-Host "Ping to $cleanPort succeeded."
                $_ | Add-Member -MemberType NoteProperty -Name Available -Value $true
            } else {
                Write-Host "Ping to $cleanPort failed."
                $ | Add-Member -MemberType NoteProperty -Name Available -Value $false
            }
        }
    }

    foreach ($device in $content) {
        if ($device.Available) {
            $onlineDevices += $device
        } else {
            $offlineDevices += $device
        }
    }

    Stop-ExcelProcess

    if ($offlineDevices.Count -gt 0) {
        $countMSG = '{0} devices found to be offline.' -f $offlineDevices.Count
        Write-Host $countMSG
        $offlineDevices | Export-Excel -Path "$env:HOMEPATH\Documents\Offline_Devices.xlsx" -WorksheetName 'Offline_Devices' -AutoSize -TableName 'Offline_Devices' -ClearSheet
    }
    if ($onlineDevices.Count -gt 0) {
        $onlineDevices | Export-Excel -Path "$env:HOMEPATH\Documents\Offline_Devices.xlsx" -WorksheetName 'Online_Devices' -AutoSize -TableName 'Online_Devices' -ClearSheet
    }
}

function Get-CommonADGroups {
    <#
    .SYNOPSIS
    Retrieves and compares Active Directory (AD) group memberships to identify groups that are common to all specified users.
    
    .DESCRIPTION
    The Get-CommonADGroups function is designed to help administrators determine which AD groups are shared among a set of users. It supports two input methods:

    A direct list of user OPIDs.
    A file containing a list of OPIDs (either a .csv or .xlsx file).
    The function queries each user's group memberships and returns only those groups that are common to all users provided.
    
    .PARAMETER Users
    - Type: String[]
    - Parameter Set: Users
    - Mandatory: Yes (when using the Users parameter set)
    - Description: A list of user OPIDs to compare.
    
    .PARAMETER UserList
    - Type: String
    - Parameter Set: UserList
    - Mandatory: Yes (when using the UserList parameter set)
    - Description: Path to a file containing user OPIDs. The file can be:
        - A .csv file with one OPID per line.
        - An .xlsx file with a column labeled OPID.    

    .EXAMPLE
    Using a list of OPIDs

    Get-CommonADGroups -Users 'opid1', 'opid2', 'opid3'

    This command compares the AD group memberships of the three specified users and returns the groups they all belong to.

    .EXAMPLE
    Using a file input

    Get-CommonADGroups -UserList 'C:\Users\Shared\userlist.xlsx'

    This command reads OPIDs from the specified Excel file and returns the groups common to all users listed.

    .OUTPUTS
    - A list of AD group names that are shared by all users.
    - Output is written to the console.
    
    .NOTES
    Requires the ActiveDirectory module.
    If using Excel input, the ImportExcel module must be available.
    Group names are extracted from the CN component of the distinguished name.
    #>

    [CmdletBinding(DefaultParameterSetName = 'Users')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Users')]
        [string[]]$Users,
        [Parameter(Mandatory = $true, ParameterSetName = 'UserList')]
        [string]$UserList
    )

    $userGroups = @{}
    if ($UserList.EndsWith('.xlsx')) { 
        $Users = Import-Excel -Path $UserList | Select-Object -ExpandProperty OPID
    } elseif ($UserList) {
        $Users = Get-Content -Path $UserList 
    }

    foreach ($user in $Users) {
        $groups = Get-ADUser -Identity $user -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        $groupNames = $groups | ForEach-Object { ($_ -split ',')[0] -replace '^CN=' }
        $userGroups[$user] = $groupNames
    }

    $commonGroups = $userGroups.Values | ForEach-Object { $_ } | Group-Object | Where-Object { $_.Count -eq $users.Count } | Select-Object -ExpandProperty Name
    Write-Output 'Groups shared by all users:'
    $commonGroups
}

function New-PopupMessage {
    <#
    .SYNOPSIS
    Sends a personalized popup message to the currently logged-in user on one or more remote computers.

    .DESCRIPTION
    The New-PopupMessage function connects to each specified computer using PowerShell remoting, retrieves the currently logged-in user's OPID (username), looks up the user's full name in Active Directory, and sends a popup message using the `msg` command. The message is personalized with the user's first and last name. Each session is removed after the message is sent.

    This function is useful for IT administrators who need to send real-time, personalized notifications to users across multiple systems.

    .PARAMETER Computers
    Specifies one or more computer names to which the popup message should be sent. The function will attempt to establish a remote session with each computer and send a message to the currently logged-in user.

    .EXAMPLE
    New-PopupMessage -Computers 'WS-12345'

    Sends a personalized popup message to the user currently logged into the computer named 'WS-12345'.

    .EXAMPLE
    New-PopupMessage -Computers 'WS-12345', 'WS-67890'

    Sends personalized popup messages to users currently logged into both 'WS-12345' and 'WS-67890'.

    .OUTPUTS
    None. The function writes status messages to the console and sends popup messages to remote users.

    .NOTES
    - Requires PowerShell remoting to be enabled on the target computers.
    - Uses the global variable `$global:cred` for authentication.
    - The message is sent using the `msg` command and includes the user's given name and surname.
    - If no user is logged in, a warning is displayed and the session is skipped.
    - Each PowerShell session is removed after use to free up resources.
    - Requires the ActiveDirectory module for user lookup.
    #>

    [CmdletBinding()]
    param (
        [string[]]$Computers,
        [string]$Message
    )

    foreach ($computer in $Computers) {
        try {
            Write-Output "Connecting to $computer..."
            $session = New-PSSession -ComputerName $computer -Credential $global:cred

            $opid = Invoke-Command -Session $session -ScriptBlock {
                (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
            }

            if (-not $opid) {
                Write-Warning "No user logged in on $computer."
                Remove-PSSession $session
                continue
            } else {
                $username = $opid.Split('\')[-1]
                $user = Get-ADUser -Identity $username -Properties GivenName, Surname
            }

            if ($user) {
                if (-not $Message) { $messageString = "Hello, $($user.GivenName) $($user.Surname)!" } else { $messageString = $Message }
                Write-Output "Sending message to $username on $($computer): $messageString"

                Invoke-Command -Session $session -ScriptBlock {
                    param ($targetUser, $msg)
                    msg $targetUser $msg
                } -ArgumentList $username, $messageString
            } else {
                Write-Warning "User $username not found in AD."
            }

            Remove-PSSession $session
        } catch {
            Write-Error "Failed to process $($computer): $_"
        }
    }
}

function Get-DuressTagUsers {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]$Path = "$env:HOMEPATH\duress_tag_data.xlsx"
    )

    $data = Import-Excel -Path $Path
    

    # Loop through each row in the imported Excel data
    $data | ForEach-Object {
        $name = $_.'Staff Name'

        # Skip if name is null or empty
        if ([string]::IsNullOrWhiteSpace($name)) {
            $_ | Add-Member -NotePropertyName 'GivenName' -NotePropertyValue $null
            $_ | Add-Member -NotePropertyName 'Surname' -NotePropertyValue $null
            return
        }

        # Remove any parenthetical nicknames or suffixes like (JR), (Kay), etc.
        $cleanedName = $name -replace '\(.*?\)', ''
        $cleanedName = $cleanedName.Trim()

        # Define the regex pattern (matches: First Last | Last, First | First Middle Last | Last, First Middle)
        $pattern = '^\s*([A-Za-z]+)\s*,?\s+([A-Za-z]+)(?:\s+([A-Za-z]\.? | [A-Za-z]+))?\s*$'

        if ($cleanedName -match $pattern) {
            # Depending on comma presence, decide field order
            if ($cleanedName -like '*,*') {
                # Format: Last, First [Middle]
                $surname = $matches[1]
                $given = $matches[2]
            } else {
                # Format: First [Middle] Last
                $given = $matches[1]
                $surname = if ($matches[3]) { $matches[3] } else { $matches[2] }
            }

            $filterString = "(GivenName -like '{0}') -and (Surname -like '{1}')" -f $given, $surname
            $adUser = Get-ADUser -Filter $filterString -Properties *
            
            if ($adUser -and ($adUser.GetType() | Where-Object { $_.Name -eq 'ADUser' })) {
                $_ | Add-Member -NotePropertyName 'OPID' -NotePropertyValue $adUser.name
                $_ | Add-Member -NotePropertyName 'Location' -NotePropertyValue $adUser.City
                $_ | Add-Member -NotePropertyName 'Enabled' -NotePropertyValue $adUser.Enabled
            } else {
                $_ | Add-Member -NotePropertyName 'OPID' -NotePropertyValue $null
                $_ | Add-Member -NotePropertyName 'Location' -NotePropertyValue $null
                $_ | Add-Member -NotePropertyName 'Enabled' -NotePropertyValue $null
            }
        } else {
            $given = $null
            $surname = $null
        }    

        # if ($given -and $surname) {
        #     [PSCustomObject]@{
        #         GivenName = $given
        #         Surname   = $surname
        #     }
        # }
    }

    Stop-ExcelProcess -Force
    $data | Export-Excel -Path "$env:HOMEPATH\sanitized_duress_tag_data.xlsx" -TableName 'Sanitized_Data' -WorksheetName 'Sanitized_Data' -ClearSheet -AutoSize
    return $data
}





function Remove-ADObjects {
    <#
    .SYNOPSIS
    Removes one or more Active Directory computer objects (optionally recursive), with full WhatIf/Confirm support.

    .DESCRIPTION
    For each provided computer identity, this function resolves the AD computer object and deletes it using
    Remove-ADObject -Recursive. It is wrapped in SupportsShouldProcess so -WhatIf/-Confirm behave as expected.
    The supplied -Credential is used for both lookup (Get-ADComputer) and deletion.

    Emits a structured result per input with Name, DistinguishedName, Removed, and Message, so you can log or export
    outcomes cleanly.

    .PARAMETER ComputerName
    Computer identity to remove. Accepts Name, DistinguishedName, GUID, or SID. Supports pipeline and
    property-based pipeline input (e.g., objects with ComputerName/Name/DNSHostName).

    .PARAMETER Credential
    Credentials to use for the AD query and deletion.

    .PARAMETER Server
    Domain controller or AD LDS instance to target.

    .PARAMETER SearchBase
    Distinguished name (DN) of the container/OU to scope the lookup.

    .PARAMETER SearchScope
    Specifies the search scope: Base, OneLevel, or Subtree.

    .INPUTS
    System.String
    You can pipe strings or objects with a ComputerName/Name/DNSHostName property.

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    Fields: Name, DistinguishedName, Removed (Boolean), Message

    .EXAMPLES
    # Remove a single computer with confirmation
    PS> Remove-ADObjects -ComputerName 'WS-123' -Credential (Get-Credential)

    # Dry-run multiple deletions
    PS> 'WS-1','WS-2','WS-3' | Remove-ADObjects -Credential (Get-Credential) -WhatIf

    # Scoped removal from a specific OU
    PS> Remove-ADObjects -ComputerName 'WS-999' -Credential (Get-Credential) `
    >> -Server 'dc01.contoso.com' -SearchBase 'OU=Workstations,DC=contoso,DC=com'
    #>

    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'DNSHostName')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNull()]
        [pscredential]$Credential,

        [Parameter()]
        [string]$Server,

        [Parameter()]
        [string]$SearchBase,

        [Parameter()]
        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [string]$SearchScope
    )

    begin {
        # Build an optional splat for scoping parameters; reused on both AD calls.
        $scope = @{}
        if ($PSBoundParameters.ContainsKey('Server')) { $scope.Server = $Server }
        if ($PSBoundParameters.ContainsKey('SearchBase')) { $scope.SearchBase = $SearchBase }
        if ($PSBoundParameters.ContainsKey('SearchScope')) { $scope.SearchScope = $SearchScope }
    }

    process {
        foreach ($comp in $ComputerName) {
            $dn = $null
            try {
                # Resolve the computer; -Identity accepts Name, DN, GUID, SID
                $adComp = Get-ADComputer -Identity $comp -Credential $Credential @scope -ErrorAction Stop
                $dn = $adComp.DistinguishedName
            } catch {
                [pscustomobject]@{
                    Name              = $comp
                    DistinguishedName = $null
                    Removed           = $false
                    Message           = "Not found or inaccessible: $($_.Exception.Message)"
                }
                continue
            }

            # Confirm deletion with ShouldProcess (honors -WhatIf / -Confirm)
            if ($PSCmdlet.ShouldProcess($dn, 'Remove AD object recursively')) {
                try {
                    # Remove-ADObject supports -Credential and -Recursive
                    Remove-ADObject -Identity $dn -Recursive -Credential $Credential @scope -ErrorAction Stop -Confirm:$false

                    [pscustomobject]@{
                        Name              = $adComp.Name
                        DistinguishedName = $dn
                        Removed           = $true
                        Message           = 'Deleted.'
                    }
                } catch {
                    [pscustomobject]@{
                        Name              = $adComp.Name
                        DistinguishedName = $dn
                        Removed           = $false
                        Message           = "Deletion failed: $($_.Exception.Message)"
                    }
                }
            } else {
                [pscustomobject]@{
                    Name              = $adComp.Name
                    DistinguishedName = $dn
                    Removed           = $false
                    Message           = 'Skipped by user (-WhatIf or -Confirm).'
                }
            }
        }
    }
}

function Repair-Printer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string[]]
        $Name
    )

    begin {
        $printerHash = @{}
        $props = @('ComputerName', 'Name', 'Location', 'Comment', 'DriverName', 'ShareName', 'Shared', 'PortName')
    }

    process {
        $printer = $mhdPrintServers | ForEach-Object -Parallel { 
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