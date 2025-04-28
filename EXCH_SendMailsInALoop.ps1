# Enter the recipient
$Recipient = "Mailbox1@itm8exchangetest.dk"

# Send 100 e-mails
1..100 | ForEach-Object {
    $Subject = "Test Email $_"
    $Body = "This is the body of test email number $_."
    Send-MailMessage -To $Recipient -Subject $Subject -Body $Body -SmtpServer "localhost" -From "Mailbox2@itm8exchangetest.dk"
    Write-Host "Sent email $_"
}

Write-Host "Created 100 emails!"
