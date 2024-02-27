<#
.DESCRIPTION
The script is a multi-tool for Exchange admins. It contains several Exchange PowerShell scripts that all can be executed by entering the corrosponding number
Save the script as "MultiTool.ps1" and run it directly inside of a elevated Exchange Shell

.SYNOPSIS
Mailbox Export
HealthCheck
Overview of Exchange configuration
Find IIS logs to a specific user or IP-address
Delete IIS logs older than 14 days
Enable or disable DAG maintenance
Show Exchange URLs
Check if the CU you have installed is supported by Microsoft
Find all GPO's that can affect Outlook behaviour

#>

# Checking permissions
$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If (-not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
Write-Host "Powershell needs to started as an Administrator" -ForeGroundColor Yellow
Break
}

Import-Module ActiveDirectory
Add-PSSnapin *EXC*

Function MultiTool { 
$Function = Read-Host "
Please enter one of the below numbers to proceed:

    1. MAILBOX EXPORT
    - Exports information about all mailboxes to a .csv-file that can be imported directly into Excel

    2. HEALTHCHECK SCRIPT
    - Checks anything from services, components to storage requirements on your Exchange server

    3. SETUP OVERVIEW
    - Exports information about your current Exchange setup in a easily readable format

    4. IIS Tracing
    - Prompts you to enter a username or IP-address you'd like to search for in all of your IIS-logs.

    5. IIS Log Cleanup
    - Deletes all IIS logfiles that is older than 14 days, to cleanup your drive

    6. Enable or disable DAG Maintenance mode
    - Helps you setting your Exchange node into Maintenance mode, or help you disable maintenance mode again

    7. SHOW EXCHANGE URLS
    - A list of all Exchange URLs

    8. Exchange CU Check
    - Will tell you if your installed Exchange CU is supported or not

    9. GPO that can affect Outlook behaviour
    - Finds all Group Policies that can affect the behaviour of Outlook

    10. EXIT
    "

    If ($Function -EQ "1")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/MailboxExportV2.ps1)"
    }

    If ($Function -EQ "2")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/HealthCheck.ps1)"
    }

    If ($Function -EQ "3")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/SetupOverview.ps1)"
    }

    If ($Function -EQ "4")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/IISUserTracing.ps1)"
    }

    If ($Function -EQ "5")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/IISLogCleanup.ps1)"
    }

    If ($Function -EQ "6")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/DAGMaintenance.ps1)"
    }

    If ($Function -EQ "7")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/ShowExchangeURLS.ps1)"
    }

    If ($Function -EQ "8")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/CUCheck.ps1)"
    }

    If ($Function -EQ "9")
    {
        cmd /c "powershell iex (irm raw.githubusercontent.com/ITR-MITHO/Microsoft-Exchange/main/GPOThatAffectOutlook.ps1)"
    }

    If ($Function -EQ "10")
    {
        Exit
    }

    
}
# Start of Function
MultiTool
