[CmdletBinding()]
param()

$opid = (whoami.exe).Replace('-a', '')

if ((-not (Get-LocalGroupMember -Group Administrators | Where-Object { $_.name -like "*$opid" }) -or (whoami.exe -ne $opid) )) {
    if ($PSCmdlet.ShouldContinue("Add $opid as a local admin?", 'You need to be a local admin running Terminal with your Domain account.')) {
        Add-LocalGroupMember -Name 'Administrators' -Member $opid
    }
    if ($PSCmdlet.ShouldContinue('Logout?', 'You need to log out and back in for local admin changes to take effect.')) {
        [int]$Time = 5
        $Length = $Time / 100
        for ($Time; $Time -gt 0; $Time--) {
            Write-Progress -Activity 'Dont forget to run this script again after logging back in...' -Status "Logging out in $Time" -PercentComplete ($Time / $Length)
            Start-Sleep 1
        }
        logoff.exe
    }
} else {
    winget install --id Microsoft.Powershell --source winget
    winget install -e --id Microsoft.VisualStudioCode
    winget install -e --id Git.Git

    # Ensure profile exists (PS7 only)
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $profilePath = $PROFILE
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


        if ($PSCmdlet.ShouldContinue('Relaunch Terminal now?', 'Terminal must be relaunched after installing PowerShell 7.')) {
            if ($PSCmdlet.ShouldContinue('Do you want to apply recommended settings to Windows Terminal?', 'You can apply optional but recommended settings automatically.')) {
                wt pwsh.exe -File "$env:OneDrive\Documents\PowerShell\terminal_setup.ps1"
                exit
            } else {
                Start-Process wt
                exit
            }
        }
    } elseif ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Host 'Make sure to use PowerShell 7 when creating your profile.' -ForegroundColor Yellow
        if ($PSCmdlet.ShouldContinue('Relaunch Terminal now?', 'Terminal must be relaunched after installing PowerShell 7.')) {
            if ($PSCmdlet.ShouldContinue('Do you want to apply recommended settings to Windows Terminal?', 'You can apply optional but recommended settings automatically.')) {
                wt pwsh.exe -File "$env:OneDrive\Documents\PowerShell\terminal_setup.ps1"
                exit
            } else {
                Start-Process wt
                exit
            }
        }
    }
}