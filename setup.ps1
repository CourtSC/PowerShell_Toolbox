[CmdletBinding()]
param()

function Start-Countdown {
    param(
        [int]$Seconds = 5,
        [string]$Message = ''
    )

    $Length = $Seconds / 100
    for ($Seconds; $Seconds -gt 0; $Seconds--) {
        Write-Progress -Activity $Message -Status $Seconds -PercentComplete ($Seconds / $Length)
        Start-Sleep 1
    }
}

$opid = (whoami.exe).Replace('-a', '')

if ((-not (Get-LocalGroupMember -Group Administrators | Where-Object { $_.name -like "*$opid" }) -or ((whoami.exe) -ne $opid) )) {
    if ($PSCmdlet.ShouldContinue("Add $opid as a local admin?", 'You need to be a local admin running Terminal with your Domain account.')) {
        Add-LocalGroupMember -Name 'Administrators' -Member $opid
    }
    if ($PSCmdlet.ShouldContinue('Logout?', 'You need to log out and back in for local admin changes to take effect.')) {
        Start-Countdown -Seconds 5 -Message 'Dont forget to run this script again after logging back in...'
        logoff.exe
        exit
    }
} else {
    winget install --id Microsoft.Powershell --source winget
    winget install -e --id Microsoft.VisualStudioCode
    winget install -e --id Git.Git

    # Ensure profile exists (PS7 only)
    $profilePath = "$env:OneDrive\Documents\PowerShell\Microsoft.PowerShell_profile.ps1"
    $profileDir = Split-Path -Parent $profilePath
    if (-not (Test-Path -Path $profilePath)) {
        New-Item -Path $profilePath -ItemType File -Force | Out-Null
    }

    Set-Location $profileDir

    # Clone repo into current directory safely
    if (-not (Test-Path '.git')) {
        git clone https://github.com/CourtSC/PowerShell_Toolbox.git temp_repo
        Copy-Item temp_repo\* -Recurse -Force
        Remove-Item temp_repo -Recurse -Force
    } else {
        git pull
    }

    # Add lines only if missing
    $importLine = 'Import-Module ITSSFunctions -Force'
    $gitPullLine = 'git -C (Split-Path -Parent $PROFILE) pull'

    $profileText = if (Test-Path $profilePath) { Get-Content -Path $profilePath -Raw } else { '' }

    if ($profileText -notmatch [regex]::Escape($gitPullLine)) {
        Add-Content -Path $profilePath -Value $gitPullLine
    }

    if ($profileText -notmatch [regex]::Escape($importLine)) {
        Add-Content -Path $profilePath -Value $importLine
    }

    Set-Location ~
}

# Configure Terminal Settings
$currSettings = Get-Content "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json" | ConvertFrom-Json

if (-not ($currSettings.profiles.list | Where-Object { $_.guid -eq '{574e775e-4f2a-5b96-ac1e-a2962a402336}' })) {
    $currSettings.profiles.list += [PSCustomObject]@{
        'commandline' = 'C:\Program Files\PowerShell\7\pwsh.exe'
        'guid'        = '{574e775e-4f2a-5b96-ac1e-a2962a402336}'
        'hidden'      = $false
        'name'        = 'PowerShell'
        'source'      = 'Windows.Terminal.PowershellCore'
    }
}

$currSettings.defaultProfile = '{574e775e-4f2a-5b96-ac1e-a2962a402336}'

if ($PSCmdlet.ShouldContinue('Do you want to apply recommended settings to Windows Terminal?', 'You can apply optional but recommended settings automatically.')) {
    # Install JetBrains Mono
    [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null
    if (-not ((New-Object System.Drawing.Text.InstalledFontCollection).Families | Where-Object { $_.Name -like '*JetBrains*' })) {
        wt winget install -e --id DEVCOM.JetBrainsMonoNerdFont
    }

    $currSettings.profiles.defaults.colorScheme = 'Campbell'
    $currSettings.profiles.defaults.font.face = 'JetBrains Mono'
    $currSettings.profiles.defaults.font.size = 16
}

$currSettings | ConvertTo-Json -Depth 100 | Set-Content -Path "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

if ($PSVersionTable.PSVersion.Major -lt 7) {
    if ($PSCmdlet.ShouldContinue('Relaunch Terminal now?', 'Terminal must be relaunched after installing PowerShell 7.')) {
        Start-Countdown -Seconds 5
        exit
    }
}

