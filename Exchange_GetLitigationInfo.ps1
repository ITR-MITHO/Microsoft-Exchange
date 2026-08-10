# Extracts info on which mailboxes has LitigationHold enabled, and how much data is in the hold

$dateStamp = Get-Date -Format "dd-MM-yyyy_HH-mm"
$OutputFile = ".\Mailbox_Litigation_Info_$dateStamp.csv"

Write-Host "Enumerating mailboxes with LitigationHold enabled. Please wait."

$LitiMailboxes = Get-Mailbox -Resultsize Unlimited | ?{$_.LitigationHoldEnabled -eq $true}
#$LitiMailboxes = Get-Mailbox -Resultsize Unlimited -RecipientTypeDetails UserMailbox

$LitiMailboxesCount = ($LitiMailboxes).count
$count = 1

Write-Host "Extracting info on mailboxes with LitigationHold enabled. Found $LitiMailboxesCount mailboxes with LitigationHold." -ForeGroundColor Green
foreach ($LitMB in $LitiMailboxes)
{  
    Write-Host "Processing mailbox $count -" $LitMB.UserPrincipalName
    #Write-Progress -Activity "Extraction in progress" -Status "Processing mailbox $count out of $LitiMailboxesCount"
    $folders = Get-MailboxFolderStatistics -Identity $LitMB.UserPrincipalName -FolderScope RecoverableItems | ?{$_.FolderPath -like "*DiscoveryHolds*" -OR $_.FolderPath -like "*Purges*"} | select FolderPath,FolderSize,ItemsInFolder
    $Results = [PSCustomObject]@{
        DisplayName = $LitMB.DisplayName
        UserPrincipalName = $LitMB.UserPrincipalName
        LitigationHoldEnabled = $LitMB.LitigationHoldEnabled
        LitigationHoldStartDate = $LitMB.LitigationHoldDate
        LitigationHoldOwner = $LitMB.LitigationHoldOwner
        LitigationHoldDuration = $LitMB.LitigationHoldDuration
        FolderPath1 = $folders[0].FolderPath
        FolderSize1 = $folders[0].FolderSize
        FolderItems1 = $folders[0].ItemsInFolder
        FolderPath2 = $folders[1].FolderPath
        FolderSize2 = $folders[1].FolderSize
        FolderItems2 = $folders[1].ItemsInFolder
        FolderPath3 = $folders[2].FolderPath
        FolderSize3 = $folders[2].FolderSize
        FolderItems3 = $folders[2].ItemsInFolder
    }

    $count++    
    $Results | Select-Object DisplayName, UserPrincipalName, LitigationHoldEnabled, LitigationHoldStartDate, LitigationHoldOwner, LitigationHoldDuration, FolderPath1, FolderSize1, FolderItems1, FolderPath2, FolderSize2, FolderItems2, FolderPath3, FolderSize3, FolderItems3 | Export-csv $OutputFile -NotypeInformation -Encoding Unicode -Delimiter "," -Append
}
Write-Host "Done. Output file found here:" -ForeGroundColor Green -NoNewline; Write-Host "$OutputFile" -ForeGroundColor Yellow

