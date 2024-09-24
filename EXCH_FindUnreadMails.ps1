$Mailbox = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox | Select Alias
$Results = @()
Foreach ($M in $Mailbox)
{

   $Search = Search-Mailbox -Identity $M.Alias -SearchQuery 'IsRead:False' -EstimateResultOnly -WarningAction SilentlyContinue | Select ResultItemsCount, ResultItemsSize
   $Results += [PSCustomObject]@{
    
    Username = $M.Alias
    Unread = $Search.ResultItemsCount
    UnreadSize = $Search.ResultItemsSize

}
    }
$Results | Select Username, Unread, UnreadSize
