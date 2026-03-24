[CmdletBinding()]
param()
$opid = Read-Host 'OPID'
if (-not (Get-LocalGroupMember -Group Administrators | Where-Object { $_.name -like "*$opid*" })) {
    Add-LocalGroupMember -Name 'Administrators' -Member $opid
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
    $wtPath = (Get-Command wt.exe).Path
    Start-Process $wtPath
    if ($PSCmdlet.ShouldContinue('Do you want to apply recommended settings to Windows Terminal?', 'You can apply optional but recommended settings automatically.')) {
        pwsh.exe -ExecutionPolicy bypass "$env:OneDrive\Documents\PowerShell\terminal_setup.ps1"
    }
    exit
}
