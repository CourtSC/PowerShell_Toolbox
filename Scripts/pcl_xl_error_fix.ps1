[CmdletBinding(DefaultParameterSetName = 'ComputerName')]
param (
    [Parameter(ParameterSetName = 'Session', Mandatory, ValueFromPipeline)]
    [ValidateNotNull()]
    [System.Management.Automation.Runspaces.PSSession]$Session,

    [Parameter(ParameterSetName = 'ComputerName', Mandatory, ValueFromPipeline)]
    [ValidateNotNullOrEmpty()]
    [string]$ComputerName,

    [Parameter(ParameterSetName = 'ComputerName')]
    [ValidateNotNullOrEmpty()]
    [pscredential]$Credential
)

$sb = {
    Stop-Service Spooler -ErrorAction SilentlyContinue

    Get-ChildItem -Path 'C:\Windows\System32\spool\drivers' -Recurse -Filter '*.gpd' | ForEach-Object { Remove-Item $_.FullName }
    Get-ChildItem -Path 'C:\Windows\System32\spool\PRINTERS' | ForEach-Object { Remove-Item $_.FullName }
    
    Restart-Service Spooler -ErrorAction SilentlyContinue
}

if ($PSCmdlet.ParameterSetName -eq 'Session') {
    Invoke-Command -Session $Session -ScriptBlock $sb -ArgumentList $PrinterName, $PrintServer, $DriverName
} elseif ($PSCmdlet.ParameterSetName -eq 'ComputerName') {
    if (-not $Credential) {
        $Credential = Get-Credential -Message 'Enter you admin (OPID-A) credentials.'
    }
    Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $sb -ArgumentList $PrinterName, $PrintServer, $DriverName
}
