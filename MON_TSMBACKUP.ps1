$Customer = "CUSTOMER NAME" # Remember to change
Add-PSSnapin *EXC*
Import-Module ActiveDirectory
$Domain = (Get-Accepteddomain | Where {$_.Default -EQ "True"}).Name
$Sender = "ITM8-EXCH@$Domain"
$Databases = Get-MailboxDatabase -Status | Select Name, LastFullBackup, LastIncrementalBackup

$FullBackupBody = ""
$IncrementalBackupBody = ""
Foreach ($DB in $Databases) {
    $Data = $DB.Name
    $LastFullBackup = $DB.LastFullBackup
    $LastIncrementalBackup = $DB.LastIncrementalBackup

    If ($LastFullBackup -lt (Get-Date).AddDays(-8)) {
        $FullBackupBody += "$Data - Last Full: $LastFullBackup`r`n"
    }

    If ($LastIncrementalBackup -lt (Get-Date).AddDays(-3)) {
        $IncrementalBackupBody += "$Data - Last Incremental: $LastIncrementalBackup`r`n"
    }
}

# Send email if there is anything to report
If ($FullBackupBody) {
    Send-MailMessage -To ExchangeTeam@itm8.com -From $Sender -Subject "$Customer - FULL BACKUP" -SmtpServer Localhost -Body "Full backups that have failed for 8 days:`r`n`r`n$FullBackupBody"
}

If ($IncrementalBackupBody) {
    Send-MailMessage -To ExchangeTeam@itm8.com -From $Sender -Subject "$Customer - INCREMENTAL BACKUP" -SmtpServer Localhost -Body "Incremental backups that have failed for 3 days:`r`n`r`n$IncrementalBackupBody"
}
