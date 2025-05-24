# Import the Active Directory module
Import-Module ActiveDirectory

# Specify the username
$username = "XXX"

# Get all domain controllers
$domainControllers = Get-ADDomainController -Filter *

# Loop through each domain controller and query the LastLogonDate for the user
foreach ($dc in $domainControllers) {
    $lastLogon = Get-ADUser -Filter {SamAccountName -eq $username} -Server $dc.HostName -Properties LastLogonDate
    if ($lastLogon.LastLogonDate) {
        Write-Host "$username is logged on to $($dc.HostName) with LastLogonDate $($lastLogon.LastLogonDate)"
    } else {
        Write-Host "$username not found on $($dc.HostName)"
    }
}