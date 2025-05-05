Add-PSSnapin *EXC*
Import-Module ActiveDirectory
$Domain = (Get-Accepteddomain | Where {$_.Default -EQ "True"}).Name
$Sender = "ITM8-EXCH@$domain"

# Export data about failed backups
If (Test-Path C:\ITM8\IncrementalLog.txt)
{
Remove-Item C:\ITM8\IncrementalLog.txt -force
Remove-Item C:\ITM8\FullLog.txt -force
}

$Databases = Get-MailboxDatabase -Status | Select Name, LastFullBackup, LastIncrementalBackup
Foreach ($DB in $Databases)
{
$Data = $DB.Name
$LastFullBackup = $DB.LastFullBackup
$LastIncrementalBackup = $DB.LastIncrementalBackup
If ($DB.LastFullBackup -LT (Get-date).AddDays(-8))
{
    Echo "$Data - Last Full: $LastFullBackup" >> C:\ITM8\FullLog.txt

Send-MailMessage -to ExchangeTeam@itm8.com -From $Sender -Subject "CUSTOMER - FULL BACKUP" -SmtpServer Localhost -Attachments C:\ITM8\FullLog.txt -Body "
Fullbackups that have failed for 8 days"
}

If ($DB.LastIncrementalBackup -LT (Get-date).AddDays(-3))
{
    Echo "$Data - Last Incremental: $LastIncrementalBackup" >> C:\ITM8\IncrementalLog.txt

Send-MailMessage -to ExchangeTeam@itm8.com -From $Sender -Subject "CUSTOMER - INCREMENTAL BACKUP" -SmtpServer Localhost -Attachments C:\ITM8\incrementalLog.txt -Body "
Incremental that have failed for 3 days"
} 
    }
