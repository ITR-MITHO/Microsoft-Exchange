$reportMailbox = "<local email address>" #Email address in the local domain
$reportMailboxFolder = "Mailbox Reports" #Folder where the reports will be placed

Write-Host "Starting search for large items in all mailboxes. Have patience, it can take a long time."

$searchresults = Get-Mailbox -ResultSize Unlimited | Search-Mailbox -SearchQuery '(size>150000000)' -EstimateResultOnly

foreach ($searchresult in $searchresults)
{
	if ($searchresult.ResultItemsCount -ne 0)
	{
		Write-Host "Large items found, so sending email report for:" $searchresult.Identity
		Search-Mailbox -Identity $searchresult.Identity -LogOnly -SearchQuery '(size>150000000)' -LogLevel Full -TargetMailbox $reportMailbox -TargetFolder $reportMailboxFolder	
	}
	elseif ($searchresult.ResultItemsCount -eq 0)
	{
		Write-Host "No large items in:" $searchresult.Identity
	}	
}
