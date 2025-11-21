if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
    Write-Host 'ImportExcel module not found. Installing...'
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    Write-Host 'ImportExcel module installed.'
} else { Import-Module ImportExcel }

function Stop-ExcelProcess {

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

$dirs = @( '\\lkmaisigrid10.multihosp.net\AH_Data\Lab\LAB_SHARED', '\\lkmaisigrid10.multihosp.net\AH_Data\facm\FACM_ENG_SHARED' )

$data = foreach ($dir in $dirs) {
    Get-ChildItem -Path $dir -Recurse -Directory | Select-Object Name, FullName, @{Name = 'Owner'; Expression = { (Get-Acl -Path $_.FullName).Owner } }
}

Stop-ExcelProcess -Force
$data | Export-Excel -Path "$env:OneDrive\Documents\H_Drive_Folder_Owners.xlsx" -WorksheetName Folder_Owners -TableName Folder_Owners -ClearSheet -AutoSize
