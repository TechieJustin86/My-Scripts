# Function to get the last logon date for a specific user across all domain controllers
function Get-UserLastLogon {
    param (
        [string]$Username
    )

    # Get all domain controllers
    $domainControllers = Get-ADDomainController -Filter *

    # Loop through each domain controller and query the LastLogonDate for the user
    foreach ($dc in $domainControllers) {
        $lastLogon = Get-ADUser -Filter {SamAccountName -eq $Username} -Server $dc.HostName -Properties LastLogonDate
        if ($lastLogon.LastLogonDate) {
            Write-Host "$Username is logged on to $($dc.HostName) with LastLogonDate $($lastLogon.LastLogonDate)"
        } else {
            Write-Host "$Username not found on $($dc.HostName)"
        }
    }
}

# Prompt for the username to search for
$username = Read-Host -Prompt "Enter the username to check"

# Call the function to get the last logon for the specified user
Get-UserLastLogon -Username $username