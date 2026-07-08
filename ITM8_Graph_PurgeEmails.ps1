Try 
{
    Connect-MgGraph -Scopes "Mail.ReadWrite" -ErrorAction Stop
}
Catch
{
    Write-Warning "MSGraph PowerShell module might be missing or authentication failed. Try: Install-Module Microsoft.Graph"
    Break
}

$CSV = Import-csv "$home\desktop\Test.csv"
foreach ($C in $CSV)
{
    $InternetID = $C.InternetMessageID
    $UserUPN    = $C.Email
    $GraphMessage = Get-MgUserMessage -UserId $UserUPN -Filter "internetMessageId eq '$InternetID'"
    if ($GraphMessage) {
        $Uri = "https://graph.microsoft.com/v1.0/users/$UserUPN/messages/$($GraphMessage.Id)/permanentDelete"
        Invoke-MgGraphRequest -Method POST -Uri $Uri -Body "{}"
        
        Write-Host "Permanently deleted message for $UserUPN" -ForegroundColor Green
    } else {
        Write-Warning "Could not find message with ID $InternetID for user $UserUPN"
    }
}
