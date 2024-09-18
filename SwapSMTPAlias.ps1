# Replaces *@primary.com with *@secondary.com as primary SMTP adress and adds *@primary.com as an alias
# TEST: $users = Get-ADUser User01 -Property ProxyAddresses

$Users = import-csv $home\desktop\Users.csv

foreach ($user in $users) {
    $proxyAddresses = $user.ProxyAddresses
    $primaryAddress = $proxyAddresses | Where-Object { $_ -like "SMTP:*@primary.com" }
    $secondaryAddress = $proxyAddresses | Where-Object { $_ -like "smtp:*@secondary.com" }

    if ($primaryAddress -and $secondaryAddress) {
        $proxyAddresses = $proxyAddresses | Where-Object { $_ -notlike "*@secondary.com" -and $_ -notlike "*@primary.com" }

        $newPrimary = $secondaryAddress -replace "smtp:", "SMTP:"  # Set *@secondary.com as primary
        $newSecondary = $primaryAddress -replace "SMTP:", "smtp:"  # Set *@primary as secondary

        $updatedProxyAddresses = @($newPrimary, $newSecondary) + $proxyAddresses | Sort-Object -Unique
        $updatedProxyAddresses = [string[]]$updatedProxyAddresses

        $UserName = $User.SamAccountName
        Set-ADUser -Identity $user -Replace @{ProxyAddresses = $updatedProxyAddresses}
        Set-ADUser -Identity $User -UserPrincipalName "$Username@secondary.com" -EmailAddress "$Username@secondary.com"
        
        Write-Host "Updated: $($user.UserPrincipalName)"
    }
}
