# ITSSFunctions.Printers Module Documentation

## Get-PrintersWithErrors
.SYNOPSIS

    Gets printers in an error or offline state and optionally removes related print jobs.

.DESCRIPTION

    Get-PrintersWithErrors queries one or more print servers for printers whose status is either Error or Offline. By default, when -Domain is set to the default value and -Servers is not specified, the function uses a built-in list of known print servers.

    If -Remove is specified, the function attempts to enumerate print jobs for each affected printer and remove them. By default, only jobs with an error-related status are removed. When -Nuke is specified, all jobs on each affected printer are targeted for removal. -Nuke implies -Remove.

    The function returns printer objects when -Remove is not used, or job-level result objects when removal is requested. Each result includes the server, printer, job ID, document name, status, and any error message encountered during processing.

    Errors encountered while querying servers or removing jobs are handled per server or per job so that processing continues for remaining printers.

.PARAMETERS
    -Domain

    DNS domain associated with the print environment. When set to the default value of multihosp.net and -Servers is not specified, the function uses the built-in default server list.

    -Servers

    One or more print servers to query.

    -Remove

    Removes matching print jobs from printers that are in an Error or Offline state.

    -Nuke

    Removes all print jobs from printers that are in an Error or Offline state. This switch implies -Remove.

.INPUTS

    None.
    You cannot pipe objects into this function.

.OUTPUTS
    System.Management.Automation.PSCustomObject

    When -Remove is not specified, outputs printer objects from Get-Printer that are in an Error or Offline state.

    System.Management.Automation.PSCustomObject

    When -Remove is specified, outputs job summary objects containing Server, Printer, JobId, Document, Status, and Message properties.

.EXAMPLES
    Example 1
    >PS Get-PrintersWithErrors
    

    Queries the default print servers for printers in an Error or Offline state and returns the matching printer objects.

    Example 2
    >PS Get-PrintersWithErrors -Servers 'PRINT01','PRINT02' -Remove
    

    Queries the specified servers, finds printers in an Error or Offline state, and removes only jobs with an error-related status.

    Example 3
    >PS Get-PrintersWithErrors -Servers 'PRINT01' -Nuke
    

    Queries the specified server, finds printers in an Error or Offline state, and removes all jobs associated with those printers.

.NOTES
    Requires the PrintManagement module.
    Uses ForEach-Object -Parallel, so PowerShell 7+ is required.
    -Nuke implies -Remove.
    Errors are written as warnings and do not stop processing other printers.

## Install-PrinterDriver
.SYNOPSIS

    Installs a selected printer driver on one or more remote computers.

.DESCRIPTION

    Install-PrinterDriver presents an interactive Windows Forms selector populated with the printer drivers currently installed on the local computer. After you choose a driver, the function validates that the selected driver exists locally, ensures the required driver files are staged on each target computer, and then installs the selected driver remotely.

    The function supports targeting computers by name or by existing PSSession objects. When -ComputerName is used, you can optionally provide -Credential. When -Session is used, -Credential is ignored. Only sessions created by this function are removed at the end of execution.

    To support larger remoting payloads, the function ensures WSMan MaxEnvelopeSizekb is at least 4096 on the local computer before any remoting begins, and on each remote computer before driver files are copied or installation starts.

    The operation honors -WhatIf and -Confirm.

    Because the driver is selected interactively, this function is intended for an interactive desktop session and is not suitable for unattended execution.

.PARAMETERS
    -ComputerName

    One or more computer names to install the selected driver on. Accepts pipeline input.

    -Session

    One or more existing PSSession objects to use. If specified, -Credential is ignored.

    -Credential

    Credentials for remoting when using -ComputerName. Ignored when -Session is used.

    -ThrottleLimit

    Maximum concurrent calls when using -ComputerName. Default is 16.

.INPUTS
    System.String

    You can pipe computer names to -ComputerName.

    Microsoft.PowerShell.Commands.PSSession

    You can pipe existing sessions to -Session.

.OUTPUTS
    System.Management.Automation.PSCustomObject

    Returns a verification object for each target computer containing ComputerName, DriverName, Installed, Version, and InfPath.

    EXAMPLES
    Example 1
    >PS 'PC01','PC02' | Install-PrinterDriver -Credential (Get-Credential) -Verbose
    

    Opens the driver selector, lets you choose a local driver, and installs that driver on the specified computers using remoting credentials.

    Example 2
    >PS $s = New-PSSession -ComputerName 'PC01','PC02'
    Install-PrinterDriver -Session $s -Confirm
    

    Uses existing sessions to install the selected local printer driver on the remote computers.

.NOTES
    Uses an interactive Windows Forms selector to choose from locally installed printer drivers.
    Requires the selected driver to already be installed or staged on the local machine.
    Requires administrative rights on the target machines.
    Uses pnputil and Add-PrinterDriver remotely.
    Ensures WSMan MaxEnvelopeSizekb is at least 4096 locally and on each target.
    PowerShell 7+ is recommended.
    Intended for interactive use only.
    LINKS
    about_Remote
    about_CommonParameters

## Set-PrinterPort

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

## Get-InstalledPrinters

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

## Get-ADPrinterGroups

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

## Get-ADPrinters

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

## Repair-Printer

_No inline documentation found._

## Set-InitialPrinterConfig

_No inline documentation found._

