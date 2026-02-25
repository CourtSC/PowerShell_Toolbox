param (
    [Parameter(ParameterSetName = 'Session', Mandatory, ValueFromPipeline)]
    [ValidateNotNull()]
    [System.Management.Automation.Runspaces.PSSession]$Session,

    [Parameter(ParameterSetName = 'ComputerName', Mandatory, ValueFromPipeline)]
    [ValidateNotNullOrEmpty()]
    [string]$ComputerName,

    [Parameter(ParameterSetName = 'ComputerName')]
    [ValidateNotNullOrEmpty()]
    [pscredential]$Credential,
    
    [Parameter(Mandatory)]
    [ValidateNotNull()]
    [string]$PrinterName,

    [Parameter(Mandatory)]
    [ValidateNotNull()]
    [string]$PrintServer,

    [Parameter(Mandatory)]
    [ValidateNotNull()]
    [string]$DriverName
)


$sb = {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param ($PrinterName, $PrintServer, $DriverName)

    Write-Warning "Performing this operation will remove ALL printers from $env:COMPUTERNAME."
    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Remove the $DriverName printer driver.")) {
        Stop-Service Spooler -ErrorAction SilentlyContinue
        Get-ChildItem -Path 'C:\Windows\System32\spool\PRINTERS' | ForEach-Object { Remove-Item $_ -Force }
        function Remove-PrinterRegKeys([string]$Sid, [string]$PrinterName) {
            $base = "Registry::HKEY_USERS\$Sid"

            $connectionsKey = "$base\Printers\Connections"
            if (Test-Path $connectionsKey) {
                Get-ChildItem $connectionsKey -ErrorAction SilentlyContinue |
                    ForEach-Object { Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }

        $profileList = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $profiles = Get-ChildItem $profileList -ErrorAction SilentlyContinue | ForEach-Object {
            $sid = $_.PSChildName
            $profilePath = (Get-ItemProperty $_.PsPath -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
            if ($profilePath) {
                [pscustomobject]@{
                    Sid       = $sid
                    NtUserDat = Join-Path $profilePath 'NTUSER.DAT'
                }
            }
        } | Where-Object { $_.Sid -match '^S-1-5-21-' -and (Test-Path $_.NtUserDat) }

        foreach ($p in $profiles) {
            $sid = $p.Sid
            $hkuPath = "Registry::HKEY_USERS\$sid"
            $alreadyLoaded = Test-Path $hkuPath
            $loadedByUs = $false

            if (-not $alreadyLoaded) {
                & reg.exe load "HKU\$sid" "$($p.NtUserDat)" | Out-Null
                $loadedByUs = Test-Path $hkuPath
            }

            if (Test-Path $hkuPath) {
                Remove-PrinterRegKeys -Sid $sid -PrinterName $PrinterName
            }

            if ($loadedByUs) {
                & reg.exe unload "HKU\$sid" | Out-Null
            }
        }

        $infPath = Get-PrinterDriver -Name $DriverName | Select-Object -ExpandProperty InfPath
        $infFileName = ($infPath -split '\\')[-1].Replace('.inf', '')
        $PrintManagementPath = 'C:\Temp\HPSPARKTools\PrintManagement\v4\HPPrintManagement.exe'
        if (-not (Test-Path -Path $PrintManagementPath)) {
            Invoke-WebRequest -Uri 'https://ftp.hp.com/pub/softlib/software13/printers/SUPD/HPSPARKTools-5.03.1.589.zip' -OutFile 'C:\Temp\HPSPARKTools.zip'
            Expand-Archive 'C:\Temp\HPSPARKTools.zip' -DestinationPath 'C:\Temp\HPSPARKTools\' -Force 
        }

        Get-ChildItem -Path 'C:\Windows\System32\spool\drivers' -Recurse -Filter $infFileName | ForEach-Object { Remove-Item $_.FullName }
        Get-ChildItem -Path 'C:\Windows\System32\spool\PRINTERS' | ForEach-Object { Remove-Item $_.FullName }
    
        Restart-Service Spooler -ErrorAction SilentlyContinue

        & $PrintManagementPath remove -d $DriverName 
    }
}

if ($PSCmdlet.ParameterSetName -eq 'Session') {
    Invoke-Command -Session $Session -ScriptBlock $sb -ArgumentList $PrinterName, $PrintServer, $DriverName
} elseif ($PSCmdlet.ParameterSetName -eq 'ComputerName') {
    if (-not $Credential) {
        $Credential = Get-Credential -Message 'Enter you admin (OPID-A) credentials.'
    }
    Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock $sb -ArgumentList $PrinterName, $PrintServer, $DriverName
}
