[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Path,
 
    [Parameter()]
    [string]$Worksheet
)

if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
    Write-Host 'ImportExcel module not found. Installing...'
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    Write-Host 'ImportExcel module installed.'
} else { Import-Module ImportExcel }

if ( $PSBoundParameters.ContainsKey('Worksheet') ) {
    $data = Import-Excel -Path $Path -WorksheetName $Worksheet
} else {
    $data = Import-Excel -Path $Path
}

$total = $data.Count
$count = 0

foreach ($printer in $data) { 
    if ($printer.PrintServerName -and $printer.'MHD Printer' -and $printer.Driver) {
        if ($PSCmdlet.ShouldProcess("\\$($printer.PrintServerName)\$($printer.'MHD Printer')", "Setting driver to $($printer.Driver)")) {
            $progress = ($count / $total) * 100
            Write-Progress -Activity "Setting driver on \\$($printer.PrintServerName)\$($printer.'MHD Printer') to $($printer.Driver)" -Status "$count/$total complete..." -PercentComplete $progress
            $count++
            Set-Printer -ComputerName $printer.PrintServerName -Name $printer.'MHD Printer' -DriverName $printer.Driver
        }
    }
}