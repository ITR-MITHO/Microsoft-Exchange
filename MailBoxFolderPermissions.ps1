param (
    [string]$Identity = "*", # Default value is "*" (all users)
    [string]$OutputFile = "$home\desktop\folderpermissions.csv"
)

# Fetch mailboxes of type UserMailbox only
$Mailboxes = Get-Mailbox -RecipientTypeDetails 'UserMailbox' $Identity -ResultSize Unlimited | Sort-Object

$result = @()

# Counter for progress bar
$MailboxCount = ($Mailboxes | Measure-Object).Count
$count = 1

foreach ($Mailbox in $Mailboxes) {

    # Use Alias property instead of name to ensure 'uniqueness' passed on to Get-MailboxFolderStatistics
    $Alias = '' + $Mailbox.Alias

    $DisplayName = ('{0} ({1})' -f $Mailbox.DisplayName, $Mailbox.Name)

    $activity = ('Working... [{0}/{1}]' -f $count, $MailboxCount)
    $status = ('Getting folders for mailbox: {0}' -f $DisplayName)
    Write-Progress -Status $status -Activity $activity -PercentComplete (($count / $MailboxCount) * 100)

    # Fetch folders
    $Folders = @('\')
    $FolderStats = Get-MailboxFolderStatistics $Alias | Select-Object -Skip 1
    foreach ($FolderStat in $FolderStats) {
        $FolderPath = $FolderStat.FolderPath.Replace('/', '\')
        $Folders += $FolderPath
    }

    foreach ($Folder in $Folders) {

        # Build folder key to fetch mailbox folder permissions
        $FolderKey = $Alias + ':' + $Folder

        # Fetch mailbox folder permissions
        $Permissions = Get-MailboxFolderPermission -Identity $FolderKey -ErrorAction SilentlyContinue

        # Store results in variable
        foreach ($Permission in $Permissions) {
            $User = $Permission.User -replace "ExchangePublishedUser\.", ""
            if ($User -EQ "Default" -and
                $Permission.AccessRights -notlike 'None') {
                $result += [PSCustomObject]@{
                    Mailbox      = $DisplayName
                    FolderName   = $Permission.FolderName
                    Identity     = $Folder
                    User         = $User -join ','
                    AccessRights = $Permission.AccessRights -join ','
                }
            }
        }
    }

    # Increment counter
    $count++
}

# Export to CSV
$result | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding Unicode
