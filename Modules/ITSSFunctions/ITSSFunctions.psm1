if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
    Write-Host 'ImportExcel module not found. Installing...'
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    Write-Host 'ImportExcel module installed.'
}

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
        [PSSession[]]$Session,

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


function Get-ADPrinterGroup {
    <#
.SYNOPSIS
Finds Active Directory groups whose Name contains a given printer name, optionally restricting to “default” groups.

.DESCRIPTION
Attempts to query with Get-ADGroup -Filter (fast, native). If the AD cmdlet throws the common
"Parameter set cannot be resolved..." error, the function automatically falls back to an LDAP
DirectorySearcher query that returns equivalent results.

.PARAMETER PrinterName
Printer name (or fragment) to search for. Single quotes in the value are safely doubled for -Filter.
For the LDAP fallback, special characters are RFC4515-escaped.

.PARAMETER Default
When specified, limits results to groups whose Name contains “DEF”.

.PARAMETER Properties
Additional AD group properties to retrieve. Defaults to 'Description'.
Use '*' for all properties (AD cmdlet path only; LDAP fallback will return a subset plus any named properties).

.PARAMETER Server
AD DC (hostname[:port]) or AD LDS instance. Used by both the AD cmdlet (when available) and the LDAP fallback.

.PARAMETER SearchBase
DN of the container/OU to search. Used by both paths.

.PARAMETER SearchScope
Base, OneLevel, or Subtree. Used by both paths.

.INPUTS
System.String

.OUTPUTS
Microsoft.ActiveDirectory.Management.ADGroup (cmdlet path)
or
System.Management.Automation.PSCustomObject (LDAP fallback)

.EXAMPLES
PS> Get-ADPrinterGroup -PrinterName 'IBJ9'
PS> Get-ADPrinterGroup -PrinterName 'HP' -Default -Properties ManagedBy,Description
PS> Get-ADPrinterGroup -PrinterName 'Xerox' -Server 'dc01.contoso.com' -SearchBase 'OU=Printers,DC=contoso,DC=com'

.NOTES
Requires RSAT ActiveDirectory module for the cmdlet path; otherwise the function will use the LDAP fallback.
#>
    [CmdletBinding()]
    [OutputType([Microsoft.ActiveDirectory.Management.ADGroup], [pscustomobject])]
    param (
        [Parameter(Position = 0, Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$PrinterName,

        [Parameter()]
        [Alias('DefaultOnly')]
        [switch]$Default,

        [Parameter()]
        [string[]]$Properties = @('Description'),

        [Parameter()]
        [string]$Server,

        [Parameter()]
        [string]$SearchBase,

        [Parameter()]
        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [string]$SearchScope
    )

    begin {
        # Map textual SearchScope to LDAP enum for fallback
        $ldapScopeMap = @{
            Base     = [System.DirectoryServices.SearchScope]::Base
            OneLevel = [System.DirectoryServices.SearchScope]::OneLevel
            Subtree  = [System.DirectoryServices.SearchScope]::Subtree
        }

        function ConvertTo-LdapEscaped {
            param([Parameter(Mandatory)][string]$InputText)
            $sb = [System.Text.StringBuilder]::new()
            foreach ($ch in $InputText.ToCharArray()) {
                switch ($ch) {
                    '(' { $null = $sb.Append('\28') }
                    ')' { $null = $sb.Append('\29') }
                    '*' { $null = $sb.Append('\2a') }
                    '\' { $null = $sb.Append('\5c') }
                    default {
                        if ([int][char]$ch -eq 0) { $null = $sb.Append('\00') }
                        else { $null = $sb.Append($ch) }
                    }
                }
            }
            $sb.ToString()
        }

        function Invoke-LdapFallback {
            param(
                [string]$PrinterName,
                [switch]$Default,
                [string]$Server,
                [string]$SearchBase,
                [string]$SearchScope,
                [string[]]$EffectiveProperties
            )

            # Build LDAP filter: (&(objectCategory=group)(name=*needle*)(name=*DEF*)?)
            $escaped = ConvertTo-LdapEscaped -InputText $PrinterName
            $nameFilter = "(name=*$escaped*)"
            if ($Default) { $nameFilter = "(&${nameFilter}(name=*DEF*))" }
            $ldapFilter = "(& (objectCategory=group) $nameFilter)".Replace(' ', '')

            # Build LDAP path
            if ([string]::IsNullOrWhiteSpace($SearchBase)) {
                $rootPath = if ($Server) { "LDAP://$Server/RootDSE" } else { 'LDAP://RootDSE' }
                $root = New-Object System.DirectoryServices.DirectoryEntry($rootPath)
                $defaultNC = $root.Properties['defaultNamingContext'][0]
                $basePath = if ($Server) { "LDAP://$Server/$defaultNC" } else { "LDAP://$defaultNC" }
            } else {
                $basePath = if ($Server) { "LDAP://$Server/$SearchBase" } else { "LDAP://$SearchBase" }
            }

            $entry = New-Object System.DirectoryServices.DirectoryEntry($basePath)
            $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
            $searcher.Filter = $ldapFilter
            $searcher.SearchScope = $ldapScopeMap[$SearchScope]
            if (-not $searcher.SearchScope) { $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree }

            # Baseline properties always returned
            $searcher.PropertiesToLoad.Clear()
            foreach ($p in @('name', 'distinguishedName', 'description')) { [void]$searcher.PropertiesToLoad.Add($p) }

            # Add any explicitly requested named props (ignore '*')
            foreach ($p in ($EffectiveProperties | Where-Object { $_ -ne '*' })) {
                [void]$searcher.PropertiesToLoad.Add($p)
            }

            $results = $searcher.FindAll()
            foreach ($r in $results) {
                $props = $r.Properties
                $obj = [ordered]@{
                    Name              = ($props['name'] | Select-Object -First 1)
                    DistinguishedName = ($props['distinguishedname'] | Select-Object -First 1)
                    ObjectClass       = 'group'
                    Description       = ($props['description'] | Select-Object -First 1)
                }
                foreach ($p in ($EffectiveProperties | Where-Object { $_ -ne '*' -and $_ -notin @('name', 'distinguishedName', 'description') })) {
                    $obj[$p] = ($props[$p] | Select-Object -First 1)
                }
                [pscustomobject]$obj
            }
        }
    }

    process {
        # ---- Normalize $Properties into a clean string[] ----------------------
        # Support callers who pass a single comma-separated string.
        if ($Properties -is [string]) {
            $Properties = $Properties -split '\s*,\s*' | Where-Object { $_ }
        }
        # Ensure it's string[] and unique (case-insensitive)
        $Properties = @($Properties | Where-Object { $_ -is [string] -and $_ } | Select-Object -Unique)

        # Default to Description if empty after normalization
        if (-not $Properties -or $Properties.Count -eq 0) {
            $Properties = @('Description')
        }

        # Build the Select-Object property list safely (no mixing raw list + variable later)
        $calcObjectClass = @{ Name = 'ObjectClass'; Expression = { $_.ObjectClass } }
        if ($Properties -contains '*') {
            $selectProps = @('*', $calcObjectClass)   # '*' already includes Name/DN/etc.
            $adProps = '*'                            # For Get-ADGroup -Properties
        } else {
            $selectProps = @('Name', 'DistinguishedName', $calcObjectClass) + $Properties
            $adProps = $Properties
        }

        # Optional params for the AD cmdlet path (only when supplied)
        $opt = @{}
        if ($PSBoundParameters.ContainsKey('Server')) { $opt.Server = $Server }
        if ($PSBoundParameters.ContainsKey('SearchBase')) { $opt.SearchBase = $SearchBase }
        if ($PSBoundParameters.ContainsKey('SearchScope')) { $opt.SearchScope = $SearchScope }

        # Safe -Filter (no outer parens; single quotes inside doubled)
        $needle = $PrinterName -replace "'", "''"
        $filter = "Name -like '*$needle*'"
        if ($Default) { $filter += " -and Name -like '*DEF*'" }

        try {
            # AD cmdlet path
            Get-ADGroup -Filter $filter -Properties $adProps @opt -ErrorAction Stop |
                Select-Object -Property $selectProps
        } catch {
            $msg = $_.Exception.Message
            if ($msg -like '*Parameter set cannot be resolved*') {
                Write-Verbose 'Get-ADGroup raised a parameter-set error. Falling back to LDAP search.'
                Invoke-LdapFallback -PrinterName $PrinterName -Default:$Default `
                    -Server $Server -SearchBase $SearchBase -SearchScope $SearchScope `
                    -EffectiveProperties $Properties |
                    Select-Object -Property $selectProps
            } else {
                Write-Error -Message "Get-ADPrinterGroup failed (Filter: $filter): $msg" -ErrorRecord $_
            }
        }
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
        [PSSession[]]$Session,

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
        [PSSession[]]$Session,

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
        [string[]]$Properties
    )

    try {
        $adObj = Get-ADComputer -Identity $Identity -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    } catch {}
    
    if (-not $adObj) {
        $adObj = Get-ADUser -Identity $Identity -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    }

    
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

function Get-PrintersWithErrors {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Domain = 'multihosp.net',

        [Parameter()]
        [string[]]$Servers
    )

    if (($Domain -eq 'multihosp.net') -and (-not $Servers)) {
        $Servers = @(
            'ADCVPRNMHDMS001', 'ADCVPRNMHDMS002', 'ADCVPRNMHDMS003', 'ADCVPRNMHDMS004',
            'ADCVPRNMHDMS005', 'ADCVPRNMHDMS006'
        )
    }

    $printers = foreach ($server in $Servers) {
        try {
            Get-Printer -ComputerName $server -ErrorAction Stop | Where-Object { $_.PrinterStatus -eq 'Error' }
        } catch {
            Write-Warning "Failed to get printers from $($server): $_"
        }
    }

    return $printers
}

function Remove-PrintJobsWithErrors {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$Domain = 'multihosp.net',

        [Parameter()]
        [string[]]$Servers,

        [Parameter()]
        [switch]$Nuke
    )

    if (($Domain -eq 'multihosp.net') -and (-not $Servers)) {
        $Servers = @(
            'ADCVPRNMHDMS001', 'ADCVPRNMHDMS002', 'ADCVPRNMHDMS003', 'ADCVPRNMHDMS004',
            'ADCVPRNMHDMS005', 'ADCVPRNMHDMS006'
        )
    }

    # Collect all printers with errors
    $printers = foreach ($server in $Servers) {
        try {
            Get-Printer -ComputerName $server -ErrorAction Stop | Where-Object { $_.PrinterStatus -eq 'Error' }
        } catch {
            Write-Warning "Failed to get printers from $($server): $_"
        }
    }

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

function Get-MHDPrinter {
    param (
        [Parameter(Mandatory, Position = 0)]
        [string[]]$Printers,
        [string[]]$Servers = (
            'ADCVPRNMHDMS001', `
                'ADCVPRNMHDMS002', `
                'ADCVPRNMHDMS003', `
                'ADCVPRNMHDMS004', `
                'ADCVPRNMHDMS005', `
                'ADCVPRNMHDMS006'
        )
    )

    foreach ($srv in $Servers) {
        foreach ($printer in $Printers) {
            try { Get-Printer -ComputerName $srv -Name $printer -ErrorAction Stop | Format-List } catch {}
        }
    }
}

function Install-RemotePrintDriver {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [System.Management.Automation.Runspaces.PSSession[]]$Session
    )

    if (-not $cred) {
        $cred = Get-Credential -Message 'Enter your admin credentials:'
    }
    if (-not $Session) {
        $computer = Read-Host -Prompt 'Computer Name'
        try {
            Write-Host "Establishing PSSession with $computer." -ForegroundColor Green
            if (Test-Connection $computer -Count 2 -Quiet) {} else { throw } 
        } catch {
            Write-Error "$computer is offline."
            return
        }
        try {
            $Session = New-PSSession -ComputerName $computer -Credential $cred -ErrorAction Stop
        } catch { 
            Write-Error "Unable to create session with $computer. Please confirm the computer name, verify it is online, and ensure you are entering your -a credentials when prompted."
            return
        }
    }
    
    # Sanity check for local driver installation
    try {
        $driverName = 'HP Universal Printing PCL 6 (v7.0.0)'
        $driverPath = Get-PrinterDriver -Name $driverName | Select-Object -ExpandProperty InfPath | Split-Path
        $driverInstalled = Test-Path $driverPath -ErrorAction Stop
        if (-not $driverInstalled) { throw }
    } catch {
        Write-Error 'Driver path not found. Please ensure driver is installed on your computer before attempting to install driver on remote computer.'
        return
    }
    
    $remoteDriverPath = $driverPath | Split-Path -Leaf
    $remotePath = 'C:\Temp\HPUPD'


    Invoke-Command -Session $Session -ScriptBlock {
        New-Item -Path 'C:\Temp\HPUPD' -ItemType Directory -Force | Out-Null
    }

    $Session | ForEach-Object {
        Write-Host "Copying files to $($_.ComputerName)." -ForegroundColor Green
        Copy-Item -Path $driverPath -Destination $remotePath -ToSession $_ -Recurse -Force | Out-Null 
    }

    Invoke-Command -Session $Session -ScriptBlock {
        param($driverName, $remotePath, $remoteDriverPath)
        $ErrorActionPreference = 'Stop'
        Write-Host "Installing driver on $env:COMPUTERNAME."
        Start-Service -Name Spooler
        pnputil.exe /add-driver ($remotePath + '\' + $remoteDriverPath + '\hpcu250u.inf') /install
        Add-PrinterDriver -Name $driverName
        Get-PrinterDriver -Name $driverName | Select-Object Name, Manufacturer, Version, InfPath
    } -ArgumentList $driverName, $remotePath, $remoteDriverPath

    foreach ($s in $Session) {
        Remove-PSSession -Session $s
    }
}