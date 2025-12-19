[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Path,
 
    [Parameter()]
    [string]$Worksheet,

    [Parameter()]
    [switch]$Color,

    [Parameter()]
    [switch]$Duplex,

    [Parameter()]
    [switch]$AutoConfirm

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
        if ($Color -and ($AutoConfirm -or $PSCmdlet.ShouldProcess(("\\$($printer.PrintServerName)\$($printer.'MHD Printer')", 'Enable color printing?')))) {
            Set-PrintConfiguration -ComputerName $printer.PrintServerName -PrinterName $printer.'MHD Printer' -Color $true
        } else {
            Set-PrintConfiguration -ComputerName $printer.PrintServerName -PrinterName $printer.'MHD Printer' -Color $false
        }
        if ($Duplex -and ($AutoConfirm -or $PSCmdlet.ShouldProcess(("\\$($printer.PrintServerName)\$($printer.'MHD Printer')", 'Enable double-sided printing?')))) {
            Set-PrintConfiguration -ComputerName $printer.PrintServerName -PrinterName $printer.'MHD Printer' -DuplexingMode TwoSidedLongEdge
        }
    }
}