$data = Import-Csv -Path 'C:\Users\sco941\Downloads\CFD SDM Resource Team Leader Assignments.csv' | Where-Object { $_.'ITSS Resource' -eq 'Court, Scott' }
$ol = New-Object -ComObject Outlook.Application
foreach ($r in $data[0]) {
    $mail = $ol.CreateItem(0); $mail.Display()
    $sig = $mail.HTMLBody
    $sig = $sig -replace '^(?:\s|(?:&nbsp;)+|(?:<br\s*/?>\s*)+|(?:<p[^>]*>\s*(?:&nbsp;|\s)*</p>\s*)+)+', ''
    $body = "<p>Hello $($r.Name),</p><p>I hope you are doing well. I am the ITSS resource assigned to work with you for the $($r.'Directory Location') H drive / shared drive review.</p><p>Please confirm the following:</p><ul><li>Are you still using data on the H drive/shared drive?</li><li>Is the data Moving or Not Moving?</li><li>What part of the H drive/shared drive is your team using?</li></ul><p>If moving, I'll follow up with next steps; if not, I need written confirmation. Please respond by Fri, 3/13/2026.</p><p>Thank you,</p>"
    $mail.HTMLBody = $body + $sig
    $mail.To = $r.'Email Address'; $mail.Subject = 'H Drive Migration'
    Start-Sleep -Milliseconds 250
}