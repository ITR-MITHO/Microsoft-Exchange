Add-PSSnapin *EXC*
Import-Module ActiveDirectory

$UserList = Get-MailBox -Resultsize Unlimited
$ExportList = @()

foreach ($User in $UserList) {

$Collection = New-Object PSObject -Property @{

FullAccess = (Get-MailboxPermission $User | Where {$_.AccessRights -EQ "FullAccess" -and -not ($_.User -like “NT AUTHORITY\*”)}).User
SendAs = (Get-Mailbox $User | Get-ADPermission | Where {$_.ExtendedRights -like "Send-As" -and -not ($_.User -like “NT AUTHORITY\*”)}).User
Name = (Get-MailboxStatistics $User).DisplayName
Size = (Get-MailboxStatistics $User).TotalItemSize.Value.ToMB()
Deleted = (Get-MailboxStatistics $User).TotalDeletedItemSize.Value.ToMB()
Total = $null
Username = (Get-Mailbox $User).SamAccountName
Email = (Get-Mailbox $User).PrimarySmtpAddress
Type = (Get-Mailbox $User).RecipientTypeDetails
DBName = (Get-MailboxStatistics $User).DatabaseName
SMTPAlias = (Get-Mailbox $User).Emailaddresses
LastLogon = (Get-MailboxStatistics $User).LastLogonTime
ADEnabled = (Get-ADUser $User.SamAccountName).Enabled

}

$ExportList += $Collection

}

# Select fields in specific order rather than random.
$ExportList | Select Username, Name, Email, Type, Size, Deleted, Total, DBName, LastLogon, ADEnabled, {$_.FullAccess}, {$_.SendAs}, {$_.SMTPAlias} | 
Export-csv C:\Users\$env:username\Desktop\ExchangeExport.csv -NoTypeInformation -Encoding Unicode




# Public Folders

Add-PSSnapin *EXC*
Import-Module ActiveDirectory

$UserList = Get-PublicFolder "\" -Recurse 
$ExportList = @()

foreach ($User in $UserList) {

$Collection = New-Object PSObject -Property @{

Name = (Get-Publicfolder $User).Name
Path = (Get-Publicfolder $User).ParentPath
MailEnabled = (Get-Publicfolder $User.ParentPath).MailEnabled
TotalItemSize = (Get-Publicfolder $User.ParentPath | Get-PublicFolderStatistics).TotalItemSize
SMTP = (Get-MailPublicFolder $User.ParentPath -ErrorAction SilentlyContinue).PrimarySMTPAddress
SMTP2 = (Get-MailPublicFolder $User -ErrorAction SilentlyContinue).PrimarySMTPAddress

}

$ExportList += $Collection

}

# Select fields in specific order rather than random.
$ExportList | Select Name, Path, MailEnabled, SMTP, SMTP2, TotalItemSize |
Export-csv C:\Users\$env:username\Desktop\PFExport.csv -NoTypeInformation -Encoding Unicode
