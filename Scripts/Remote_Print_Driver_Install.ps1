function Install-RemotePrintDriver {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [System.Management.Automation.Runspaces.PSSession[]]$Session,
        [Parameter()]
        [pscredential]$Credential
    )

    if (-not $Credential) {
        $Credential = Get-Credential -Message 'Enter your admin credentials:'
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
            $Session = New-PSSession -ComputerName $computer -Credential $Credential -ErrorAction Stop
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
Clear-Host
Install-RemotePrintDriver
