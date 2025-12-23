[CmdletBinding()]
param (
    [Parameter(Position = 0, Mandatory)]
    [string]
    $Server,

    [Parameter(Position = 1, Mandatory)]
    [string]
    $PrinterName
)

begin {
    $printProperties = @{
        'Config:DuplexUnit'         = 'Installed'
        'Config:JobStorageControl'  = 'auto'
        'Config:AccessoryOutputBin' = 'NoOutputBin'
        'Config:SecurePrintControl' = 'auto'
        'Config:TintTestingControl' = 'disable'
    }
}

process {
    Set-PrintConfiguration -ComputerName $Server -PrinterName $PrinterName -DuplexingMode OneSided -PaperSize Letter -Collate $true
    $printProperties.GetEnumerator() | ForEach-Object {
        Set-PrinterProperty -ComputerName $Server -PrinterName $PrinterName -PropertyName $_.Name -Value $_.Value
    }
}