<#

.DESCRIPTION
- Veeam Full Backup will alert if the backup is over 24-hours old.
- Incrementalbackup will alert if the backup is over 3 days old.
- DifferentialBackup will alert if the backup is over 3 days old.
- Servers with incremental or differential will also trigger an alert if the fullbackup is older than 8-days. 

#>

<# Variables #>
Add-PSSnapin *EXC* 
$Timestamp = Get-MailboxDatabase -Status
$Date = (Get-date).AddDays(-1) # Veeam Full Snapshot Backup
$Date3 = (Get-date).AddDays(-3) # Incremental & Differential
$Date8 = (Get-date).AddDays(-8) # Servers with incremental and differential
 
<# Veeam Full Snapshot Backup #>

If (-not $Timestamp.LastDifferentialBackup -and -not $Timestamp.LastIncrementalBackup)
    {
    If ($Timestamp.LastFullBackup -LT $Date)
{
    Write-Host "Veeam full backup is over 24-hours old!" -ForegroundColor Red
}
else 
{
    Write-Host "Veaam full OK"
}
    }
 
<# Incremental Backup #>
If (-not $Timestamp.LastDifferentialBackup)
{
    If (-not $Timestamp.LastIncrementalBackup)
    {

    }
    else    
{
        
If ($Timestamp.LastIncrementalBackup -LT $Date3)
 
{
Write-Host "Incremental backup is over 3 days old!" -ForegroundColor Red
}
        }
            }
 
<# Differential Backup #>
If (-not $Timestamp.LastIncrementalBackup)
 
{
    If (-not $Timestamp.LastDifferentialBackup)
    {

    }
 
Else 
   
{
 
If ($Timestamp.LastDifferentialBackup -LT $Date3)
    
{
Write-Host "Differential backup is over 3 days old!" -ForegroundColor Red
}
        }
            }
 
<# 8-day Fullbackup #>
 
If ($Timestamp.LastFullBackup -LT $Date8)
{
Write-Host "Full backup is over 8 days old!" -ForegroundColor Red        
}
else 
{    
Write-Host "8-day full backup is OK"   
}
