<#
    This script is used to check if the Service Control Manager (scmanager) is using the default set of permissions.
    If default, will modify permissions to add the user defined SID.

    This was written to puppetize permissions on hosts to enable automation of services (start/stop/pausing)
    using a service account with no admin rights.

    Configuration:
    1. Define the USER or GROUP SID variable. SID can be found by
        - Get-ADgroup -Identity <AD Group> | select SID
        - Get-ADuser -Identity <AD user> | select SID

    To revert permission changes made by this script, there are two options:
        - Run "sc sdset scmanager "insert default SDDL string"
        - Delete "Security" key in the registery located in HKLM\SYSTEM\CurrentControlSet\Control\ServiceGroupOrder\Security, then reboot the system.

    References:
    - http://woshub.com/granting-remote-access-on-scmanager-to-non-admin-users/
    - http://woshub.com/set-permissions-on-windows-service/
#>

# Group or user SID
$SID = "<insert SID Here>"

# Default Permissions in SDDL format
$DefaultPermissions = "D:(A;;CC;;;AU)(A;;CCLCRPRC;;;IU)(A;;CCLCRPRC;;;SU)(A;;CCLCRPWPRC;;;SY)(A;;KA;;;BA)(A;;CC;;;AC)S:(AU;FA;KA;;;WD)(AU;OIIOFA;GA;;;WD)"

# Query service manager to get the current permissions in SDDL format.
$CheckPermissionsCommand = @'
cmd.exe /C  sc sdshow scmanager
'@

$SetPermissionsCommand = @'
cmd.exe /C  sc sdset scmanager "D:(A;;CC;;;AU)(A;;CCLCRPRC;;;IU)(A;;CCLCRPRC;;;SU)(A;;CCLCRPWPRC;;;SY)(A;;KA;;;BA)(A;;CC;;;AC)(A;;CCLCRPRC;;;$($SID))S:(AU;FA;KA;;;WD)(AU;OIIOFA;GA;;;WD)"
'@

$HostPermissions = Invoke-Expression -Command:$CheckPermissionsCommand


# Test if default permissions are in place.
If ($HostPermissions -eq $DefaultPermissions) {
    Write-Host -ForegroundColor Yellow "Found default scmanager permissions. Adding $($SID)"
    Invoke-Expression -Command:$SetPermissionsCommand
    write-host "New permission set: "
    Invoke-Expression -Command:$CheckPermissionsCommand
}
