py.exe 'C:\Users\sco941\OneDrive - AdventHealth\Documents\Python_CSV_Splitter\split_by_first_subdir.py'

$files = Get-ChildItem -Path 'C:\Users\sco941\OneDrive - AdventHealth\Documents\H_Drive_Scans\Full_Scans\Python' | Select-Object -ExpandProperty FullName

foreach ($file in $files) {
    $fileName = ($file -split '__')[-1]
    Write-Host "Processing $fileName"
    $content = Import-Csv -Path $file | Where-Object { (($_.Files -ne 0) -or ($_.Folders -ne 0)) -and ($_.Name -notmatch $pattern) }
    $root = $content[0].Name
    $content | ForEach-Object {
        $depth = ($($_.Name).Replace("$root", '') -split '\\' | Where-Object { $_.Trim() -ne '' }).count
        $_ | Add-Member -NotePropertyName Depth -NotePropertyValue $depth -Force
    }
    $content | Export-Csv -Path "$outpath\Full_Scans\$fileName" -Force
}