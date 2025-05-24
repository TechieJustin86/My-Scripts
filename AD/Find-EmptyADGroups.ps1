<#
This script is useful for administrators who want to identify and manage groups in Active Directory that have no members. 
Empty groups may be unused or need to be cleaned up, and this script helps in quickly finding them for further review or action.
#>
# Import the Active Directory module
Import-Module ActiveDirectory

# Get all groups in the domain
$groups = Get-ADGroup -Filter *

# Initialize an array to store groups with no members
$emptyGroups = @()

# Loop through each group and check for members
foreach ($group in $groups) {
    # Get the group members
    $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue

    # Check if the group has no members
    if ($members.Count -eq 0) {
        $emptyGroups += $group
    }
}

# Output the groups with no members
if ($emptyGroups.Count -gt 0) {
    Write-Host "The following groups have no members:" -ForegroundColor Green
    $emptyGroups | ForEach-Object {
        Write-Host $_.Name -ForegroundColor Yellow
    }
} else {
    Write-Host "No empty groups found in the domain." -ForegroundColor Green
}