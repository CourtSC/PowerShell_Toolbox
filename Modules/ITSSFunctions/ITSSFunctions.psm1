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
        [System.Management.Automation.PSCredential]$Credential,

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

        [Parameter(Mandatory = $true, ParameterSetName = 'ComputerName')]
        [System.Management.Automation.PSCredential]$Credential,

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
            Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $scriptBlock -ArgumentList $Filter
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
    Optional credential for Invoke-Command.

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
        [System.Management.Automation.PSCredential]$Credential
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
    Retrieves the Active directory groups that a specified user or computer is a member of.

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
        [System.Management.Automation.PSCredential]$Credential,

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

function Get-DirTreeSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [switch]$Recurse,

        [string]$ExportPath = "$env:OneDrive\Documents\DirectoryTreeSize.csv",

        # Optional: don’t scan deeper than this (huge perf win on monster shares)
        [int]$MaxDepth = [int]::MaxValue,

        # Optional: skip owner lookups (ACL calls are slow on shares)
        [switch]$SkipOwner
    )

    $ErrorActionPreference = 'Stop'
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $rootPath = (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')

    if (-not $Recurse) {
        $files = Get-ChildItem -LiteralPath $rootPath -File -ErrorAction Stop
        $fileStats = $files | Measure-Object -Property Length -Sum
        $directoryCount = (Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction Stop | Measure-Object).Count

        $lastWrite = if ($files.Count) { ($files | Measure-Object LastWriteTime -Maximum).Maximum } else { $null }
        $lastAccess = if ($files.Count) { ($files | Measure-Object LastAccessTime -Maximum).Maximum } else { $null }

        [pscustomobject]@{
            Path                = $rootPath
            Parent              = $null
            Depth               = 0
            Owner               = if ($SkipOwner) { $null } else { (Get-Acl -LiteralPath $rootPath).Owner }
            FileCount           = $fileStats.Count
            DirectoryCount      = $directoryCount
            DirSizeInMB         = if ($fileStats.Sum) { [math]::Round($fileStats.Sum / 1MB, 3) } else { 0 }
            TotalSizeInMB       = if ($fileStats.Sum) { [math]::Round($fileStats.Sum / 1MB, 3) } else { 0 }
            LastModified        = if ($lastWrite) { $lastWrite.ToString('MM/dd/yyyy HH:mm') } else { $null }
            LastAccessed        = if ($lastAccess) { $lastAccess.ToString('MM/dd/yyyy HH:mm') } else { $null }
            TotalRuntimeSeconds = $null
        }
        return
    }

    # --- Aggregation tables (keyed by normalized dir path) ---
    $sizeBytesByDir = @{}
    $fileCountByDir = @{}
    $lastWriteByDir = @{}
    $lastAccessByDir = @{}
    $dirSet = [System.Collections.Generic.HashSet[string]]::new()


    # Ensure root exists in tables
    $dirSet.Add($rootPath) | Out-Null
    $sizeBytesByDir[$rootPath] = 0L
    $fileCountByDir[$rootPath] = 0

    # --- Stream files once ---
    Get-ChildItem -LiteralPath $rootPath -File -Recurse -Force -ErrorAction Stop | ForEach-Object {
        $dir = $_.DirectoryName.TrimEnd('\')

        # Depth filter (optional)
        if ($MaxDepth -ne [int]::MaxValue) {
            $rel = $dir.Substring($rootPath.Length).TrimStart('\')
            $depth = if ($rel) { ($rel -split '\\').Count } else { 0 }
            if ($depth -gt $MaxDepth) { return }
        }

        $dirSet.Add($dir) | Out-Null

        if (-not $sizeBytesByDir.ContainsKey($dir)) {
            $sizeBytesByDir[$dir] = 0L
            $fileCountByDir[$dir] = 0
        }

        $sizeBytesByDir[$dir] += [int64]$_.Length
        $fileCountByDir[$dir] += 1

        $lw = $_.LastWriteTime
        if (-not $lastWriteByDir.ContainsKey($dir) -or $lw -gt $lastWriteByDir[$dir]) {
            $lastWriteByDir[$dir] = $lw
        }

        $la = $_.LastAccessTime
        if (-not $lastAccessByDir.ContainsKey($dir) -or $la -gt $lastAccessByDir[$dir]) {
            $lastAccessByDir[$dir] = $la
        }
    }

    # --- Enumerate directories (also streaming) ---
    Get-ChildItem -LiteralPath $rootPath -Directory -Recurse -Force -ErrorAction Stop | ForEach-Object {
        $dir = $_.FullName.TrimEnd('\')

        if ($MaxDepth -ne [int]::MaxValue) {
            $rel = $dir.Substring($rootPath.Length).TrimStart('\')
            $depth = if ($rel) { ($rel -split '\\').Count } else { 0 }
            if ($depth -gt $MaxDepth) { return }
        }

        $dirSet.Add($dir) | Out-Null
        if (-not $sizeBytesByDir.ContainsKey($dir)) {
            $sizeBytesByDir[$dir] = 0L
            $fileCountByDir[$dir] = 0
        }
    }

    # Build list of dirs and compute parent/depth
    # Build list of directories in a robust way
    if ($null -eq $dirSet) {
        throw 'Unexpected: dirSet is null'
    }

    # If dirSe# Build list of directories in a robust way
    if ($null -eq $dirSet) {
        throw 'Unexpected: dirSet is null'
    }

    if ($dirSet -is [System.Array]) {
        $allDirs = $dirSet
    } elseif ($dirSet -is [System.Collections.IEnumerable]) {
        try {
            $allDirs = $dirSet.ToArray()
        } catch {
            $allDirs = @($dirSet)
        }

        # Normalize to strings
        $allDirs = $allDirs | ForEach-Object { $_.ToString() }
    } else {
        $allDirs = @($dirSet.ToString())
    }
    $rows = New-Object System.Collections.Generic.List[object]

    # Child mapping for totals and directory counts
    $children = @{}
    foreach ($dir in $allDirs) {
        if ($dir -eq $rootPath) { continue }
        $parent = (Split-Path -Path $dir -Parent).TrimEnd('\')
        if (-not $children.ContainsKey($parent)) {
            $children[$parent] = New-Object System.Collections.Generic.List[string]
        }
        $children[$parent].Add($dir) | Out-Null
    }

    foreach ($dir in $allDirs) {
        $parent = if ($dir -eq $rootPath) { $null } else { (Split-Path -Path $dir -Parent).TrimEnd('\') }

        $rel = if ($dir -eq $rootPath) { '' } else { $dir.Substring($rootPath.Length).TrimStart('\') }
        $depth = if ($rel) { ($rel -split '\\').Count } else { 0 }

        $dirBytes = [int64]$sizeBytesByDir[$dir]
        $dirMB = [math]::Round($dirBytes / 1MB, 3)

        $owner = $null
        if (-not $SkipOwner) {
            try { $owner = (Get-Acl -LiteralPath $dir -ErrorAction Stop).Owner } catch { $owner = $null }
        }

        $rows.Add([pscustomobject]@{
                Path           = $dir
                Parent         = $parent
                Depth          = $depth
                Owner          = $owner
                FileCount      = [int]$fileCountByDir[$dir]
                DirectoryCount = if ($children.ContainsKey($dir)) { $children[$dir].Count } else { 0 }
                DirSizeInMB    = $dirMB
                TotalSizeInMB  = 0.0  # filled later
                LastModified   = if ($lastWriteByDir.ContainsKey($dir)) { $lastWriteByDir[$dir].ToString('MM/dd/yyyy HH:mm') } else { $null }
                LastAccessed   = if ($lastAccessByDir.ContainsKey($dir)) { $lastAccessByDir[$dir].ToString('MM/dd/yyyy HH:mm') } else { $null }
            }) | Out-Null
    }

    # --- Compute TotalSize bottom-up (directories only) ---
    $byPath = @{}
    foreach ($r in $rows) { $byPath[$r.Path] = $r }

    foreach ($r in ($rows | Sort-Object Depth -Descending)) {
        $total = [double]$r.DirSizeInMB
        if ($children.ContainsKey($r.Path)) {
            foreach ($c in $children[$r.Path]) {
                $total += [double]$byPath[$c].TotalSizeInMB
            }
        }
        $r.TotalSizeInMB = [math]::Round($total, 3)
    }

    $sw.Stop()
    $rootRow = $byPath[$rootPath]
    $rootRow | Add-Member -NotePropertyName TotalRuntimeSeconds -NotePropertyValue ([math]::Round($sw.Elapsed.TotalSeconds, 3)) -Force

    # --- CSV export (fast, opens clean in Excel) ---
    $rows |
        Sort-Object Depth, Path |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

    Write-Verbose ('Exported {0} rows to {1}. Runtime: {2}' -f $rows.Count, $ExportPath, $sw.Elapsed)
    $rows
}
 