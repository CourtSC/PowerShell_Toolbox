# ITSSFunctions Module Reference

This document describes the functions in `ITSSFunctions.psm1` in a PowerShell help style format, adapted for a GitHub README.

## Requirements

Some functions in this module depend on external modules or environmental setup:

- **ImportExcel** for Excel import/export operations.
- **ActiveDirectory** for `Get-ADUser`, `Get-ADGroup`, `Get-ADComputer`, and related cmdlets.
- **PowerShell remoting** for remote queries and remote actions.
- **Windows desktop Excel** for functions that stop Excel processes.
- Several functions accept optional `-Credential` parameters for remoting or Active Directory operations.

---

## Stop-ExcelProcess

### SYNOPSIS
Stops any running Microsoft Excel processes, with optional confirmation and `WhatIf` support.

### DESCRIPTION
`Stop-ExcelProcess` looks for running `EXCEL` processes and stops them if the user confirms the action. The function supports standard PowerShell common parameters through `SupportsShouldProcess`, so `-WhatIf` and `-Confirm` work as expected.

By default, the function prompts for confirmation with `ShouldContinue`. Use `-Force` to skip that extra prompt.

### PARAMETERS

#### -Force
Skips the interactive confirmation prompt and attempts to close Excel immediately. The function still respects `-WhatIf` and `-Confirm`.

### INPUTS
None. You cannot pipe input to this function.

### OUTPUTS
`System.Boolean`

Returns `$true` when Excel is already closed or was successfully stopped. Returns `$false` when the operation is cancelled or an error occurs.

### EXAMPLES
```powershell
Stop-ExcelProcess
```
Prompts before closing any running Excel processes.

```powershell
Stop-ExcelProcess -Force
```
Closes Excel immediately without the extra confirmation prompt.

```powershell
Stop-ExcelProcess -WhatIf
```
Shows what would happen without making any changes.

### NOTES
- Uses `Get-Process -Name 'EXCEL'` and `Stop-Process -Force`.
- Requires Windows desktop Excel.
- Supports `-WhatIf` and `-Confirm`.

---

## Get-Locked

### SYNOPSIS
Checks whether one or more remote computers are locked.

### DESCRIPTION
`Get-Locked` determines lock state by checking for the `LogonUI` process on the target computer. It can query either by computer name or by existing `PSSession` objects.

When `-Monitor` is specified, the function repeatedly checks until the computer becomes locked or the optional timeout expires.

### PARAMETERS

#### -ComputerName
One or more remote computer names to query through PowerShell remoting. Accepts pipeline input.

#### -Session
One or more existing `PSSession` objects to query.

#### -Credential
Credentials used when connecting to remote computers by name.

#### -Monitor
Keeps checking until the computer becomes locked or until the timeout is reached.

#### -IntervalSeconds
Polling interval used with `-Monitor`. Default is `5` seconds.

#### -TimeoutSeconds
Maximum amount of time to wait when using `-Monitor`. Default is `0`, which means no timeout.

#### -ThrottleLimit
Maximum number of concurrent remote calls when using `-ComputerName`. Default is `16`.

### INPUTS
`System.String`, `System.Management.Automation.Runspaces.PSSession`

You can pipe computer names or `PSSession` objects.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

Returned objects include:

- `ComputerName`
- `IsLocked`
- `Timestamp`

### EXAMPLES
```powershell
Get-Locked -ComputerName WS-12345
```
Checks whether `WS-12345` is currently locked.

```powershell
Get-Locked -ComputerName WS-12345 -Monitor -TimeoutSeconds 300
```
Waits up to five minutes for the computer to become locked.

```powershell
Get-Locked -Session $session
```
Checks lock state through an existing session.

### NOTES
- Uses `Get-Process -Name 'LogonUI'` on the remote computer.
- Supports both direct computer-name queries and existing sessions.

---

## Get-InstalledSoftware

### SYNOPSIS
Retrieves installed software from remote computers.

### DESCRIPTION
`Get-InstalledSoftware` reads the uninstall registry keys on a remote computer and returns software entries from both the 64-bit and 32-bit uninstall locations.

If a software entry was installed through MSI, the function attempts to return the product GUID as well.

### PARAMETERS

#### -ComputerName
One or more remote computers to query.

#### -Session
One or more existing `PSSession` objects to use instead of computer names.

#### -Filter
Optional partial-name filter. Only software whose display name matches the filter is returned.

### INPUTS
None directly.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

Each returned object includes:

- `Name`
- `Version`
- `Publisher`
- `InstallDate`
- `UninstallString`
- `GUID`
- `RegistryPath`

### EXAMPLES
```powershell
Get-InstalledSoftware -ComputerName WS-12345
```
Returns installed software on the remote computer.

```powershell
Get-InstalledSoftware -ComputerName WS-12345 -Filter Microsoft
```
Returns only software entries whose names contain `Microsoft`.

```powershell
Get-InstalledSoftware -Session $session
```
Uses an existing remote session.

### NOTES
- Queries `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall` and `HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`.
- The function uses the provided `-Credential` value when connecting by computer name.
- Requires registry access on the remote computer.

---

## Remove-Software

### SYNOPSIS
Uninstalls software on a remote computer.

### DESCRIPTION
`Remove-Software` performs a silent uninstall by using either an MSI product code GUID or a full uninstall command string.

The uninstall is executed remotely and the function returns a structured status object that includes the log path created during the operation.

### PARAMETERS

#### -ComputerName
One or more target computers to uninstall software from.

#### -Session
One or more existing `PSSession` objects.

#### -GUID
The MSI product code to uninstall. Must be in the form `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.

#### -UninstallString
A complete uninstall command line for non-MSI uninstallers.

#### -Credential
Optional credential used for `Invoke-Command`.

### INPUTS
None directly.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

The returned object includes:

- `ComputerName`
- `Identifier`
- `Status`
- `LogPath`
- `Error` when the uninstall fails

### EXAMPLES
```powershell
Remove-Software -ComputerName WS-12345 -GUID '{12345678-1234-1234-1234-1234567890AB}'
```
Uninstalls an MSI package by product code.

```powershell
Remove-Software -ComputerName WS-12345 -UninstallString '"C:\Program Files\App\uninstall.exe" /quiet'
```
Runs a silent uninstall command.

### NOTES
- Requires either `-GUID` or `-UninstallString`.
- Uses `msiexec.exe /x` for MSI removals.
- Writes a transcript log under `C:\Windows\Temp` on the remote computer.
- Uses `Start-Transcript` and `Stop-Transcript` around the uninstall.

---

## Get-EODTime

### SYNOPSIS
Calculates a projected end-of-day time.

### DESCRIPTION
`Get-EODTime` estimates when the current work period will end based on hours already worked and a start time.

If `HoursWorked` is 40 or less, the function calculates the time required to reach 40 hours. If more than 40 hours have been worked, it calculates the time required to reach 80 hours. A 30-minute buffer is always added.

### PARAMETERS

#### -HoursWorked
The number of hours already worked.

#### -StartTime
The starting time used as the base for the calculation. The value must be accepted by `Get-Date`.

### INPUTS
None directly.

### OUTPUTS
`System.DateTime`

The function writes a calculated date and time value.

### EXAMPLES
```powershell
Get-EODTime -HoursWorked 32 -StartTime '2025-07-28 08:00'
```
Calculates when 40 hours will be reached, starting from 8:00 AM.

```powershell
Get-EODTime -HoursWorked 45 -StartTime '2025-07-28 09:00'
```
Calculates when 80 hours will be reached, starting from 9:00 AM.

### NOTES
- Adds a 0.5-hour buffer to the result.
- Useful for shift planning and time tracking.

---

## Get-ADGroupMemberships

### SYNOPSIS
Retrieves the Active Directory groups that a user or computer belongs to.

### DESCRIPTION
`Get-ADGroupMemberships` first attempts to resolve the supplied identity as a computer account. If that fails, it attempts to resolve it as a user account.

The function then retrieves each group in the `MemberOf` list and optionally returns selected group properties.

### PARAMETERS

#### -Identity
The user or computer identity to query.

#### -Properties
Optional list of group properties to return. If omitted, the default group object properties are returned.

### INPUTS
`System.String`

### OUTPUTS
`Microsoft.ActiveDirectory.Management.ADGroup`-derived objects

When `-Properties` is supplied, the function returns selected group properties. Otherwise, it returns the default group object output.

### EXAMPLES
```powershell
Get-ADGroupMemberships -Identity jdoe
```
Returns all groups for the user `jdoe`.

```powershell
Get-ADGroupMemberships -Identity srv-web01$ -Properties Name,Description
```
Returns the `Name` and `Description` properties for groups associated with the computer account.

### NOTES
- Requires the ActiveDirectory module.
- Writes a heading to the console before returning the group objects.
- Attempts `Get-ADComputer` first, then `Get-ADUser`.

---

## Get-ADGroupMembers

### SYNOPSIS
Retrieves members of Active Directory groups.

### DESCRIPTION
`Get-ADGroupMembers` can collect members of groups using one of three input methods:

- a list of group names,
- a file containing group names,
- or a directory structure whose folder names are used to infer group names.

The function can optionally exclude matching group names and can export results to Excel when used with file or directory input.

### PARAMETERS

#### -GroupNames
One or more group name patterns to search for. This parameter belongs to the `GroupNames` parameter set.

#### -GroupListPath
Path to a file containing group names, one per line. This parameter belongs to the `GroupListPath` parameter set.

#### -DirectoryPath
Path to a directory structure used to infer group names from subfolder names. This parameter belongs to the `DirectoryPath` parameter set.

#### -Exclude
One or more patterns to exclude from the group search.

### INPUTS
None directly.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

For standard group lookups, the function returns objects with group and member display name information. When `-DirectoryPath` or `-GroupListPath` is used, the function also exports results to Excel.

### EXAMPLES
```powershell
Get-ADGroupMembers -GroupNames 'IT-Admin', 'HR-Users'
```
Returns members of groups matching the supplied names.

```powershell
Get-ADGroupMembers -GroupListPath 'C:\Groups\groupnames.txt' -Exclude 'Test', 'Temp'
```
Reads group names from a file and excludes matches containing `Test` or `Temp`.

```powershell
Get-ADGroupMembers -DirectoryPath 'C:\DepartmentFolders'
```
Infers group names from folder structure and exports the results to Excel.

### NOTES
- Requires the ActiveDirectory module.
- Requires the ImportExcel module for Excel export.
- Uses the `multihosp.net` domain in AD queries.
- Output is written to `$env:HOMEPATH\Documents\AD Groups`.

---

## Get-OfflineDevices

### SYNOPSIS
Checks whether devices listed in an Excel file are online or offline.

### DESCRIPTION
`Get-OfflineDevices` imports an Excel file, reads the `Port` column, cleans trailing underscore suffixes, and pings each entry to determine whether it is reachable.

The function then separates devices into online and offline collections and exports the results to `Offline_Devices.xlsx` in the user's Documents folder.

### PARAMETERS

#### -Path
Path to the Excel file that contains the device list.

### INPUTS
None directly.

### OUTPUTS
None explicitly returned.

The function writes the results to an Excel file and reports counts to the console.

### EXAMPLES
```powershell
Get-OfflineDevices -Path 'C:\Network\devices.xlsx'
```
Checks the devices in the workbook and exports the results.

### NOTES
- Requires the ImportExcel module.
- Expects a `Port` column in the workbook.
- Uses `Test-Connection` to test reachability.
- Calls `Stop-ExcelProcess` before exporting.
- Writes output to `$env:HOMEPATH\Documents\Offline_Devices.xlsx`.

---

## Get-CommonADGroups

### SYNOPSIS
Finds Active Directory groups common to all specified users.

### DESCRIPTION
`Get-CommonADGroups` compares the `MemberOf` memberships of multiple users and returns only the groups that appear for every user in the input set.

The function accepts either a direct list of OPIDs or a file containing OPIDs. If the file ends in `.xlsx`, the function reads the `OPID` column.

### PARAMETERS

#### -Users
One or more user OPIDs to compare.

#### -UserList
Path to a file containing user OPIDs.

### INPUTS
None directly.

### OUTPUTS
`System.String`

The function writes a label line and then outputs the names of the groups shared by all specified users.

### EXAMPLES
```powershell
Get-CommonADGroups -Users 'opid1', 'opid2', 'opid3'
```
Returns the groups shared by the listed users.

```powershell
Get-CommonADGroups -UserList 'C:\Users\Shared\userlist.xlsx'
```
Reads OPIDs from an Excel file and returns the groups common to all users listed.

### NOTES
- Requires the ActiveDirectory module.
- Requires the ImportExcel module when using `.xlsx` input.
- Group names are derived from the CN portion of each distinguished name.

---

## New-PopupMessage

### SYNOPSIS
Sends a personalized popup message to the currently logged-in user on one or more remote computers.

### DESCRIPTION
`New-PopupMessage` connects to each specified computer, finds the currently logged-in user, looks up that user in Active Directory, and sends a popup message with the `msg` command.

If no custom message is supplied, the function builds a greeting from the user's first and last name.

### PARAMETERS

#### -Computers
One or more computer names to target.

#### -Message
Optional custom popup text. If omitted, a default greeting is used.

### INPUTS
None directly.

### OUTPUTS
None.

The function writes status messages and warnings to the console.

### EXAMPLES
```powershell
New-PopupMessage -Computers 'WS-12345'
```
Sends a personalized popup message to the user currently logged into the target computer.

```powershell
New-PopupMessage -Computers 'WS-12345', 'WS-67890' -Message 'Please save your work and reboot.'
```
Sends the same custom message to both remote computers.

### NOTES
- Requires PowerShell remoting to be enabled.
- Uses the provided `-Credential` parameter for authentication.
- Requires the ActiveDirectory module.
- Sends the message with the Windows `msg` command.

---

## Get-DuressTagUsers

### SYNOPSIS
Matches staff names in an Excel file to Active Directory users.

### DESCRIPTION
`Get-DuressTagUsers` imports an Excel workbook, reads the `Staff Name` column, and attempts to split each name into given name and surname.

The function then searches Active Directory for a matching user and adds the following fields to each row when a match is found:

- `OPID`
- `Location`
- `Enabled`

The cleaned data is then exported to a sanitized Excel file.

### PARAMETERS

#### -Path
Path to the input Excel file. The default is `$env:HOMEPATH\duress_tag_data.xlsx`.

### INPUTS
None directly.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

The function returns the imported rows after adding the enrichment columns.

### EXAMPLES
```powershell
Get-DuressTagUsers
```
Uses the default input file.

```powershell
Get-DuressTagUsers -Path 'C:\Data\duress_tag_data.xlsx'
```
Processes a specific workbook.

### NOTES
- Requires the ImportExcel module.
- Requires the ActiveDirectory module.
- Expects a `Staff Name` column in the workbook.
- Stops Excel before exporting the cleaned workbook.
- Writes the output to `sanitized_duress_tag_data.xlsx` in the user profile folder.

---

## Remove-ADObjects

### SYNOPSIS
Removes one or more Active Directory computer objects.

### DESCRIPTION
`Remove-ADObjects` resolves each computer identity and removes the corresponding AD computer object recursively.

The function supports `-WhatIf` and `-Confirm` through `SupportsShouldProcess` and returns a structured result for each input item.

### PARAMETERS

#### -ComputerName
The computer identity to remove. Accepts `Name`, `DistinguishedName`, `GUID`, or `SID` and supports pipeline input.

#### -Credential
Credentials used for AD lookup and deletion.

#### -Server
Optional domain controller or AD LDS instance to target.

#### -SearchBase
Optional DN used to scope the search.

#### -SearchScope
Optional search scope. Valid values are `Base`, `OneLevel`, and `Subtree`.

### INPUTS
`System.String`

You can pipe strings or objects with a `ComputerName`, `Name`, or `DNSHostName` property.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

Returned objects include:

- `Name`
- `DistinguishedName`
- `Removed`
- `Message`

### EXAMPLES
```powershell
Remove-ADObjects -ComputerName 'WS-123' -Credential (Get-Credential)
```
Removes a single computer object after confirmation.

```powershell
'WS-1','WS-2','WS-3' | Remove-ADObjects -Credential (Get-Credential) -WhatIf
```
Shows what would be removed without making changes.

```powershell
Remove-ADObjects -ComputerName 'WS-999' -Credential (Get-Credential) -Server 'dc01.contoso.com' -SearchBase 'OU=Workstations,DC=contoso,DC=com'
```
Scopes the lookup to a specific OU and domain controller.

### NOTES
- Requires the ActiveDirectory module.
- Uses `Remove-ADObject -Recursive`.
- Returns a result object even when the object is not found or the deletion is skipped.

---

## Get-DirTreeSize

### SYNOPSIS
Calculates the size of a directory tree and exports the results to CSV.

### DESCRIPTION
`Get-DirTreeSize` analyzes a directory tree and returns size, file count, directory count, owner, and date information for each folder.

Without `-Recurse`, it reports the root directory only. With `-Recurse`, it walks the tree, computes per-directory sizes, calculates total sizes bottom-up, and exports the results to CSV.

### PARAMETERS

#### -Path
Root folder to analyze.

#### -Recurse
Scans the full tree below the root path.

#### -ExportPath
Path to the CSV file that will receive the exported report. The default is `$env:OneDrive\Documents\DirectoryTreeSize.csv`.

#### -MaxDepth
Optional maximum directory depth to scan.

#### -SkipOwner
Skips owner lookups for faster execution.

### INPUTS
None directly.

### OUTPUTS
`System.Management.Automation.PSCustomObject`

Returned objects include:

- `Path`
- `Parent`
- `Depth`
- `Owner`
- `FileCount`
- `DirectoryCount`
- `DirSizeInMB`
- `TotalSizeInMB`
- `LastModified`
- `LastAccessed`
- `TotalRuntimeSeconds` on the root row

### EXAMPLES
```powershell
Get-DirTreeSize -Path 'C:\Data'
```
Reports the top-level folder summary only.

```powershell
Get-DirTreeSize -Path 'C:\Data' -Recurse
```
Scans the entire tree and exports the full report.

```powershell
Get-DirTreeSize -Path '\\Server\Share' -Recurse -MaxDepth 3 -SkipOwner
```
Scans only three levels deep and skips owner lookups.

### NOTES
- Exports a CSV report after processing.
- Designed for large folder trees where performance matters.
- Uses `Get-Acl` for owner lookup unless `-SkipOwner` is supplied.

---

## Module Summary

This module is geared toward IT support and desktop administration tasks, especially:

- remote workstation checks,
- Active Directory lookups and cleanup,
- software inventory and removal,
- Excel-based reporting,
- and file system reporting.

