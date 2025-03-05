# Import Active Directory module
Import-Module ActiveDirectory

# Define variables
$Manager = "SAMACCOUNTNAME"  # Replace with the actual manager's SAMAccountName
$ManagerDN = (Get-ADUser -Identity $Manager).DistinguishedName
$Groups = import-csv $home\desktop\groups.csv 

foreach ($Group in $Groups) {
$GroupName = $Group.SamAccountName
    $GroupDN = (Get-ADGroup -Identity $GroupName).DistinguishedName
    Set-DistributionGroup -Identity $GroupName -ManagedBy $Manager

    $ACL = Get-ACL -Path ("AD:\$GroupDN")
    $Identity = New-Object System.Security.Principal.NTAccount($Manager)
    $ADRights = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
    $Type = [System.Security.AccessControl.AccessControlType]::Allow
    $AccessRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($Identity, $ADRights, $Type, [Guid]"bf9679c0-0de6-11d0-a285-00aa003049e2")
    $ACL.AddAccessRule($AccessRule)
    Set-ACL -Path ("AD:\$GroupDN") -AclObject $ACL

    Write-Output "Manager $Manager has been granted permission to update membership for $GroupName."
}
