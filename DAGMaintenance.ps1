<#
.SYNOPSIS
Based on your input the script will either put Exchange into maintenance mode, or take it out of it. 
Everything is done automatically, and works by simply inputting a number into the promt when asked. 

.NOTES
* MAKE SURE TO RUN THIS ON THE SERVER YOU WANT TO PUT INTO MAINTENANCE MODE


#>

# Checking permissions

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
Write-Host "Start PowerShell as an Administrator" -ForeGroundColor Yellow
Break
}

Import-Module ActiveDirectory
Add-PSSnapin *EXC*

Function DAG { 
    $Function = Read-Host "
Please enter one of the below numbers to proceed:

    1 - START MAINTENANCE MODE
    2 - STOP MAINTENANCE MODE
    3 - CHECK MAINTENANCE STATE
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

If ($Function -EQ "3")
{
$HubTransport = Get-ServerComponentState -Identity "$env:computername" -Component HubTransport | Select State
If ($HubTransport.State -eq "Active")
{
    Write-Host "HubTransport is ACTIVE" -ForegroundColor Green
}
Else
{
    Write-Host "HubTransport is STOPPED" -ForegroundColor Yellow
}

$ServerWideOffline = Get-ServerComponentState "$env:computername" -Component ServerWideOffline | Select State
If ($ServerWideOffline.State -eq "Active")
{
    Write-Host "ServerWideOffline is ACTIVE" -ForegroundColor Green
}
Else
{
    Write-Host "ServerWideOffline is STOPPED" -ForegroundColor Yellow
}

$ClusterNode = Get-ClusterNode $env:computername | Select State
If ($ClusterNode.State -EQ "Up")
{
    Write-Host "ClusterNode is ACTIVE" -ForegroundColor Green
}
Else
{
    Write-Host "ClusterNode is STOPPED" -ForegroundColor Yellow
}

$DatabaseCopy = Get-MailboxServer "$env:computername"  | Select DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow
If ($DatabaseCopy.DatabaseCopyAutoActivationPolicy -EQ "Unrestricted" -and $DatabaseCopy.DatabaseCopyActivationDisabledAndMoveNow -eq $false)
{
    Write-Host "DatabaseCopy is ACTIVE" -ForegroundColor Green
}
Else
{
    Write-Host "DatabaseCopy is STOPPED" -ForegroundColor Yellow
}

Write-Host "
INFORMATION: Ensure that databases is redistributed correctly, if they are not distributed correctly follow AliTajrans steps: 
https://www.alitajran.com/balance-mailbox-databases-in-exchange-dag/#h-balance-exchange-mailbox-database-with-powershell-script

Listed below is the activation preferenses and where they are mounted currently." -ForegroundColor Yellow

Get-MailboxDatabase | ft Name, ActivationPreference, Server -AutoSize
}

# This needs to be here to start the loop.
DAG
