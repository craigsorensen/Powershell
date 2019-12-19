<#
    This script is used to check if the Service Control Manager (scmanager) is using the default set of permissions.
    If default, will return True, if not default, will return False.

    This was written to puppetize permissions on hosts to enable automation of services (start/stop/pausing)
    using a service account with no admin rights.
#>

# Default Permissions in SDDL format
$DefaultPermissions = "D:(A;;CC;;;AU)(A;;CCLCRPRC;;;IU)(A;;CCLCRPRC;;;SU)(A;;CCLCRPWPRC;;;SY)(A;;KA;;;BA)(A;;CC;;;AC)S:(AU;FA;KA;;;WD)(AU;OIIOFA;GA;;;WD)"

# Query service manager to get the current permissions in SDDL format.
$command = @'
cmd.exe /C  sc sdshow scmanager
'@

$HostPermissions = Invoke-Expression -Command:$command


# Test if default permissions are in place.
If ($HostPermissions -eq $DefaultPermissions) {
    return $true
}

return $false
