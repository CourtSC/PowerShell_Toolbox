[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Path
)

$data = Import-Csv -Path $Path | Where-Object { ($_.'ITSS Resource' -eq 'Court, Scott') -and ($_.'Project Status (For AIT)' -eq 'Not Started') }

if (-not $data -or $data.Count -eq 0) {
    Write-Host "No rows found for 'Court, Scott'."
    return
}

$nextFri = ((Get-Date).AddDays(([DayOfWeek]::Friday - $date.DayOfWeek + 7) % 7)).ToString('MM/dd/yyyy')

# Try to get/create Outlook COM object with retries
function Get-OutlookApplication {
    param($tries = 3, $waitSeconds = 2)
    for ($i = 1; $i -le $tries; $i++) {
        try {
            Stop-Process -Name OUTLOOK -ErrorAction SilentlyContinue
            $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop
            return $ol
        } catch {
            Write-Host "Attempt $($i): Could not create Outlook COM object: $($_.Exception.Message)"
            # If Outlook process not running, try to start it
            if (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
                Write-Host 'Starting Outlook...'
                Start-Process -FilePath 'outlook' -ErrorAction SilentlyContinue
                Start-Sleep -Seconds $waitSeconds
            } else {
                Start-Sleep -Seconds $waitSeconds
            }
        }
    }
    return $null
}

$Outlook = Get-OutlookApplication -tries 5 -waitSeconds 3

if (-not $Outlook) {
    Write-Error "Failed to create Outlook COM object. Ensure Outlook is installed, that you're running PowerShell with the same privilege as Outlook, and try again."
    return
}

# For each row create & display a mail item
foreach ($row in $data) {
    try {
        $body = @"
<p>Hello $($row.Name),</p>

<p>
I hope you are doing well. I am the ITSS resource assigned to work with you for the 
$($row.'Directory Location') H drive / shared drive review.
</p>

<p>Please confirm the following:</p>

<ul>
<li>Are you still using data on the H drive/shared drive?</li>
<li>Is the data Moving or Not Moving?</li>
<li>What part of the H drive/shared drive is your team using?</li>
</ul>

<p>
If the data is moving, I will follow up with the next steps/template. If it is not moving, 
I will need your written confirmation for documentation.
</p>

<p>
If possible, please send your response by Friday, $nextFri.
</p>

<p>
Thank you,<br><br>
Scott C. Court<br>
AdventHealth<br>
IT Support Specialist - Senior | Orlando
</p>
"@


        $Mail = $Outlook.CreateItem(0) # 0 = MailItem
        $Mail.To = $row.'Email Address'
        $Mail.Subject = 'H Drive Migration'
        $Mail.HTMLBody = $body
        # $Mail.Body = $body
        $Mail.Display()  # use .Send() to send immediately
    } catch {
        Write-Warning "Failed to create/display mail for $($row.'Email Address'): $($_.Exception.Message)"
    }
}