<# 
 
.DESCRIPTION  
This script gives X (user or group) access to calendars for X users.
Default is all users, the accessrights can be changed under variables, so can the user/group.
 
 
.OUTPUTS 
It has a log, other than that, it just changes calendar permissions.
It's got a test function, if you remove the #/Hashtag from line 114 (the one that changes permissions)
 
 it outputs the file to "C:\office365Scripts\Keys\Credentials.txt"
If the file is not present, one will be made.
#>
 
 
#Variables
$username                   = admin@YourTenant.onmicrosoft.com
[string]$Log                = "C:\Office365Scripts\Calendar.log"
[string]$UserToGiveAccess   ="default" #If the setting is org-wide, write "Default"
[string]$AccessRight        ="reviewer"
$pathoffile                 = "C:\office365Scripts\Keys"
$filename                   = "credentials.txt"
 

 

#PasswordCheck
$totalpath                  = "$pathoffile\$filename"
if (!(Test-Path $totalpath)){
    New-Item -ItemType directory $pathoffile
    New-Item -path $pathoffile -Name "Credentials.txt" -ItemType file
    Write-Host "Created new file"
    $password = read-host "Enter the password of the user"
    $secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
    $secureStringText = $secureStringPwd | ConvertFrom-SecureString 
 
    Set-Content $totalpath $secureStringText
    }
else
{
  Write-Host
}
 
Start-Transcript -Path $Log -Force
 

 

#login
$password = Get-Content $totalpath | ConvertTo-SecureString
$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
$session = New-PSSession –ConnectionUri https://ps.outlook.com/powershell –AllowRedirection –Authentication Basic –Credential $psCred –ConfigurationName Microsoft.Exchange
$Import = Import-Pssession $Session -AllowClobber
Import-Module MSOnline -verbose
 

 

#################################################################################################################
##############################################START OF ACTUAL SCRIPT#############################################
#################################################################################################################
 
 
foreach($mbx in Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox){
    $languageCalendar = (Get-MailboxFolderStatistics -Identity $mbx.userprincipalname -FolderScope Calendar | Select-Object -first 1).name
    Set-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User $UserToGiveAccess -AccessRights $AccessRight
    Get-MailboxFolderPermission -Identity ($mbx.UserPrincipalName+":\$LanguageCalendar") -User $UserToGiveAccess | Select-Object identity,accessrights
}
