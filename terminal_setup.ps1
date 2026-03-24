# Optional but recommended configs for Windows Terminal
[CmdletBinding()]
param()

# Install JetBrains Mono
if (-not ((New-Object System.Drawing.Text.InstalledFontCollection).Families | Where-Object { $_.Name -like '*JetBrains*' })) {
    winget install -e --id DEVCOM.JetBrainsMonoNerdFont
}

# Configure Terminal Settings
$currSettings = Get-Content "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json" | ConvertFrom-Json
$pwshGuid = $currSettings.profiles.list | Where-Object { $_.commandline -like '*pwsh.exe' } | Select-Object -ExpandProperty guid

if ( $pwshGuid ) {
    $currSettings.defaultProfile = $pwshGuid
    $currSettings.profiles.defaults.colorScheme = 'Campbell'
    $currSettings.profiles.defaults.font.face = 'JetBrains Mono'
    $currSettings.profiles.defaults.font.size = 16
    $currSettings | ConvertTo-Json -Depth 100 | Set-Content -Path "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

} else {
    if ($PSCmdlet.ShouldContinue('Relaunch Terminal now?', 'Terminal must be relaunched after installing PowerShell 7.')) {
        $wtPath = (Get-Command wt.exe).Path
        Start-Process $wtPath
        pwsh.exe -ExecutionPolicy bypass "$env:OneDrive\Documents\PowerShell\terminal_setup.ps1"
        exit
    }
}