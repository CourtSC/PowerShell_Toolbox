$servers = @(
    'FHOSVMWPRN001', 'FHOSVMWPRN002', 'FHOSVMWPRN003', 'FHOSVMWPRN004', 'FHOSVMWPRN005', 'FHOSVMWPRN006',
    'FHOVPRNA001', 'FHOVPRNB001', 'FHOVPRNC001', 'FHOVPRND001', 'FHOVPRNE001', 'FHOVPRNF001'
)

$pattern = '(?i)\\\\(' + ($servers -join '|') + ')(\\|$)'

function Process-UserHive([string]$Sid) {
    $base = "Registry::HKEY_USERS\$Sid"

    $connectionsKey = "$base\Printers\Connections"
    if (Test-Path $connectionsKey) {
        Get-ChildItem $connectionsKey -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $pattern } |
            ForEach-Object { Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue }
    }

    $devicesPath = "$base\Software\Microsoft\Windows NT\CurrentVersion\Devices"
    $printerPortsPath = "$base\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"

    foreach ($vp in @($devicesPath, $printerPortsPath)) {
        if (Test-Path $vp) {
            $item = Get-Item $vp -ErrorAction SilentlyContinue
            foreach ($name in $item.GetValueNames()) {
                $data = [string]$item.GetValue($name)
                if ($name -match $pattern -or $data -match $pattern) {
                    Remove-ItemProperty -Path $vp -Name $name -ErrorAction SilentlyContinue
                }
            }
        }
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
        Process-UserHive -Sid $sid
    }

    if ($loadedByUs) {
        & reg.exe unload "HKU\$sid" | Out-Null
    }
}

$livePrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -match $pattern -or ($_.PortName -match '(?i)\\\\(' + ($servers -join '|') + ')\\')
}

foreach ($prn in $livePrinters) {
    # Try native removal first
    try {
        Remove-Printer -Name $prn.Name -ErrorAction Stop
    } catch {
        # Fallback: PrintUI (works well on stubborn connections)
        & rundll32.exe printui.dll, PrintUIEntry /dn /n "$($prn.Name)" | Out-Null
        & rundll32.exe printui.dll, PrintUIEntry /gd /n "$($prn.Name)" | Out-Null
    }
}

Restart-Service Spooler -ErrorAction SilentlyContinue
