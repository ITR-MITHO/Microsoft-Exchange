<#
.SYNOPSIS
Based on your input the script will either put Exchange into maintenance mode, or take it out of it. 
Everything is done automatically, and works by simply inputting a number into the promt when asked. 

.NOTES
* MAKE SURE TO RUN THIS ON THE SERVER YOU WANT TO PUT INTO MAINTENANCE MODE


#>

# Checking permissions
$PMError = Test-Path $Home\desktop\PermissionIssue.txt
if ($PMError)
{
Remove-Item "$Home\desktop\PermissionIssue.txt" -Force
}
timeout 3
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
echo "Start PowerShell as an Administrator" > $Home\desktop\PermissionIssue.txt
Start $home\desktop\PermissionIssue.txt
Break
}

Import-Module ActiveDirectory
Add-PSSnapin *EXC*

Function DAG { 
    $Function = Read-Host "
Please enter one of the below numbers to proceed:

    1 - START MAINTENANCE MODE
    2 - STOP MAINTENANCE MODE
    0 - EXIT
    "
    
    If ($Function -EQ "1")
    {

Write-Host "

IMPORTANT - PLEASE READ:
If you choose to continue, $env:computername will be placed into maintenance mode. This will cause the server to stop responding to any incomming connections.

" -ForegroundColor Yellow
$Confirm = Read-Host "Are you sure you want to continue? (Y/N)"
If ($Confirm -eq "Y")

{

    $Domain = (Get-ADDomain).DNSRoot
    $Redirect = Get-ExchangeServer | Where Name -NotLike "*$env:computername*" | Select Name -First 1
    $RedirectName = $Redirect.Name
    
    
    Write-Host "INFORMATION: Putting $env:computername into maintenance mode" -ForeGroundColor Yellow
    Set-ServerComponentState -Identity "$env:computername" -Component HubTransport -State Draining -Requester Maintenance
    Redirect-Message -Server "$env:computername" -Target "$RedirectName.$Domain" -Confirm:$false
    Move-ActiveMailboxDatabase -Server $env:computername -SkipMoveSuppressionChecks -Confirm:$false | out-null
    Write-Host "

...Moving mounted databases to another node with healthy copies

" -ForegroundColor Yellow

    Timeout 15 | Out-Null
    Suspend-ClusterNode "$env:computername"
    Set-MailboxServer "$env:computername" -DatabaseCopyActivationDisabledAndMoveNow $true -DatabaseCopyAutoActivationPolicy Blocked
    Set-ServerComponentState "$env:computername" -Component ServerWideOffline -State Inactive -Requester Maintenance
    Clear
    Write-Host "$env:computername is now in maintenance mode." -ForegroundColor Green
       
}
Else 

    {
    
DAG

    }
    
}

    If ($Function -EQ "2")
    {
        Write-Host "INFORMATION: Taking $env:computername out of maintenance mode" -ForegroundColor yellow
        $DAG = (Get-DatabaseavailabilityGroup | Where Servers -like "*$env:computername*" | Select Name).Name
        Set-ServerComponentState "$env:computername" -Component ServerWideOffline -State Active -Requester Maintenance
        Resume-ClusterNode -Name "$env:computername"
        Set-MailboxServer "$env:computername" -DatabaseCopyAutoActivationPolicy Unrestricted -DatabaseCopyActivationDisabledAndMoveNow $false
        Set-ServerComponentState "$env:computername" -Component HubTransport -State Active -Requester Maintenance

        Timeout 10 | Out-Null
        # Redistribute databases
        CD "$env:exchangeinstallpath\Scripts"
        .\RedistributeActiveDatabases.ps1 -DagName "$DAG" -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks -Confirm:$false
        
 
        Write-Host "INFORMATION: $env:computername is now online and databases have been distributed." -ForegroundColor Green
    }

    If ($Function -EQ "0")
    {
        Exit
}
    }


# This needs to be here to start the loop.
DAG
