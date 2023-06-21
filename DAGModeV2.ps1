<#
.SYNOPSIS
The script is designed to either put or take an Exchange Server in/out of Maintenence mode.

#>

Function DAG {
    CLS    
    $Function = Read-Host "
    
    Enter 1 - Start Maintenence mode
    Enter 2 - Stop Maintenence mode
    Enter 0 - Exit
    "
    
    If ($Function -EQ "1")
    {

Write-Host "

IMPORTANT - PLEASE READ:
If you choose to continue, $env:computername will be placed into maintenence mode. This will cause the server to stop responding to any incomming connections.

" -ForegroundColor Yellow
$Confirm = Read-Host "Are you sure you want to continue? (Y/N)"
If ($Confirm -eq "Y")

{

    $Domain = (Get-ADDomain).DNSRoot
    $Redirect = Get-ExchangeServer | Where Name -NotLike "*$env:computername*" | Select Name -First 1
    
    
    Write-Host "INFORMATION: Putting $env:computername.$Domain into maintenence mode" -ForeGroundColor Yellow
    Set-ServerComponentState -Identity "$env:computername.$Domain" -Component HubTransport -State Draining -Requester Maintenance
    cls
  
  
    Redirect-Message -Server "$env:computername.$Domain" -Target "$Redirect" -Confirm:$false
    Suspend-ClusterNode "$env:computername.$Domain"
    Set-MailboxServer "$env:computername.$Domain" -DatabaseCopyActivationDisabledAndMoveNow $true
    Set-MailboxServer "$env:computername.$Domain" -DatabaseCopyAutoActivationPolicy Blocked
    Set-ServerComponentState "$env:computername.$Domain" -Component ServerWideOffline -State Inactive -Requester Maintenance
    Timeout 15 | Out-Null
    
    Clear
    Write-Host "$env:computername.$Domain is now in maintenence mode." -ForegroundColor Green
       
}
Else 

    {
    
DAG

    }
    
}

    If ($Function -EQ "2")
    {
        Write-Host "INFORMATION: Taking $env:computername.$Domain out of maintenence mode" -ForegroundColor yellow
        $DAG = (Get-DatabaseavailabilityGroup | Where Servers -like "*$env:computername*" | Select Name).Name
        Set-ServerComponentState "$env:computername.$Domain" -Component ServerWideOffline -State Active -Requester Maintenance
        Resume-ClusterNode -Name "$env:computername.$Domain"
        Set-MailboxServer "$env:computername.$Domain" -DatabaseCopyAutoActivationPolicy Unrestricted
        Set-MailboxServer "$env:computername.$Domain" -DatabaseCopyActivationDisabledAndMoveNow $false
        Set-ServerComponentState "$env:computername.$Domain" -Component HubTransport -State Active -Requester Maintenance
        Timeout 15 | Out-Null

        # Redistribute databases
        cd $exscripts
        .\RedistributeActiveDatabases.ps1 -DagName "$DAG" -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks -Confirm:$false
        
        Clear
        Write-Host "INFORMATION: $env:computername.$Domain is now online and databases have been distributed." -ForegroundColor Green
    }

    If ($Function -EQ "0")
    {
        Exit
}
    }


# This needs to be here to start the loop.
DAG
