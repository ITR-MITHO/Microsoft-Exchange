<#

.DESCRIPTION
The script is designed to gather information about all ActiveSync devices in an on-prem Exchange & Exchange Online environment. 
It can be run without editing, but needs to be running in an elevated PowerShell

.NOTES
If you have any issues, contact me directly for guidance at mitho@itrelation.dk

.OUTPUTS
1 file named "EASDevices.csv" will be created and placed on your desktop.

#>

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
write-host "Script is not running as Administrator" -ForegroundColor Yellow
Break
}

Add-PSSnapin *EXC*
$Date = Get-Date
$report = @()

$Stats = @("DeviceID",
            "DeviceAccessState",
            "DeviceModel"
            "DeviceType",
            "DeviceFriendlyName",
            "DeviceOS",
            "LastSyncAttemptTime",
            "LastSuccessSync"
          )


$Export = "$home\desktop\EASDevices.csv"
Write-Host "Starting to analyze.. This could take a while depending on the size of your organisation. Grab a coffee :)" -ForegroundColor Yellow
$MailboxesWithEASDevices = @(Get-CASMailbox -Resultsize Unlimited | Where {$_.HasActiveSyncDevicePartnership})

Foreach ($Mailbox in $MailboxesWithEASDevices)
{
    
    $EASDeviceStats = @(Get-ActiveSyncDeviceStatistics -Mailbox $Mailbox.Identity -WarningAction SilentlyContinue)
    $MailboxInfo = Get-Mailbox $Mailbox.Identity | Select DisplayName,PrimarySMTPAddress
    
    Foreach ($EASDevice in $EASDeviceStats)
    { 
        $lastsyncattempt = ($EASDevice.LastSyncAttemptTime)

        if ($lastsyncattempt -eq $null)
        {
            $syncAge = "Never"
        }
        else
        {
            $syncAge = ($Date - $lastsyncattempt).Days
        }

        if ($syncAge -ge $Age -or $syncAge -eq "Never")

        {

            $reportObj = New-Object PSObject
            $reportObj | Add-Member NoteProperty -Name "Display Name" -Value $MailboxInfo.DisplayName
            $reportObj | Add-Member NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
            $reportObj | Add-Member NoteProperty -Name "Sync Age (Days)" -Value $syncAge
                
            Foreach ($stat in $stats)
            {
                $reportObj | Add-Member NoteProperty -Name $stat -Value $EASDevice.$stat
            }

            $report += $reportObj
        }
    }
}
cls
Write-Host "Report completed. Find your file here: $Export" -ForegroundColor Gree
$report | Export-Csv $Export -NoTypeInformation  -Encoding UTF8
