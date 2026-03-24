[CmdletBinding()]
param()
$opid = whoami.exe
if (-not (Get-LocalGroupMember -Group Administrators | Where-Object { $_.name -like "*$opid" })) {
    if ($PSCmdlet.ShouldContinue("Add $opid as a local admin?", 'You need to be a local admin running Terminal with your Domain account.')) {
        Add-LocalGroupMember -Name 'Administrators' -Member $opid
    }
    if ($PSCmdlet.ShouldContinue('Logout?', 'You need to log out and back in for local admin changes to take effect.')) {
        logoff.exe
    }
}

winget install --id Microsoft.Powershell --source winget
winget install -e --id Microsoft.VisualStudioCode
winget install -e --id Git.Git

pwsh -NoProfile -Command "
`$profilePath = `$PROFILE
`$profileDir  = Split-Path -Parent `$profilePath

# Ensure profile exists (PS7 only)
if ((-not (Test-Path -Path `$profilePath)) -and (`$PSVersionTable.PSVersion.Major -ge 7)) {
    New-Item -Path `$profilePath -ItemType File -Force | Out-Null
} elseif (`$PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'Make sure to use PowerShell 7 when creating your profile.' -ForegroundColor Yellow
}

Set-Location `$profileDir

# Clone repo into current directory safely
if (-not (Test-Path '.git')) {
    git clone https://github.com/CourtSC/PowerShell_Toolbox.git temp_repo
    Copy-Item temp_repo\* -Recurse -Force
    Remove-Item temp_repo -Recurse -Force
} else {
    git pull
}

# Add lines only if missing
`$importLine  = 'Import-Module ITSSFunctions -Force'
`$gitPullLine = 'git -C (Split-Path -Parent `$PROFILE) pull'

`$profileText = if (Test-Path `$profilePath) { Get-Content -Path `$profilePath -Raw } else { '' }

if (`$profileText -notmatch [regex]::Escape(`$gitPullLine)) {
    Add-Content -Path `$profilePath -Value `$gitPullLine
}

if (`$profileText -notmatch [regex]::Escape(`$importLine)) {
    Add-Content -Path `$profilePath -Value `$importLine
}

Set-Location ~
"

if ($PSCmdlet.ShouldContinue('Relaunch Terminal now?', 'Terminal must be relaunched after installing PowerShell 7.')) {
    if ($PSCmdlet.ShouldContinue('Do you want to apply recommended settings to Windows Terminal?', 'You can apply optional but recommended settings automatically.')) {
        wt pwsh.exe -NoExit -ExecutionPolicy bypass "$env:OneDrive\Documents\PowerShell\terminal_setup.ps1"
    } else {
        Start-Process wt
    }
    exit
}
