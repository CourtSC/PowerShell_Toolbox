
Stop-Service Spooler -ErrorAction SilentlyContinue

Get-ChildItem -Path 'C:\Windows\System32\spool\drivers' -Recurse -Filter '*.gpd' | ForEach-Object { Remove-Item $_.FullName }
Get-ChildItem -Path 'C:\Windows\System32\spool\PRINTERS' | ForEach-Object { Remove-Item $_.FullName }
    
Restart-Service Spooler -ErrorAction SilentlyContinue
Restart-Computer -Force
