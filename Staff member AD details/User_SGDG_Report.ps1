Search all users in AD
# Working 1/9/25
# Import the Active Directory module
Import-Module ActiveDirectory

# Prompt for the first and last name
$firstName = Read-Host -Prompt "Enter the user's first name"
$lastName = Read-Host -Prompt "Enter the user's last name"

# Check if both first and last name were entered
if (-not $firstName -or -not $lastName) {
    Write-Host "Both first and last name are required. Exiting script." -ForegroundColor Red
    exit
}

# File export location with dynamic naming based on first and last name
$csvPath = "C:\Temp\$firstName-$lastName-GroupsReport.csv"

# Find the user in Active Directory by first and last name
$user = Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Property DistinguishedName, MemberOf

# Check if the user was found
if (-not $user) {
    Write-Host "No user found with the name $firstName $lastName. Exiting script." -ForegroundColor Red
    exit
}

# Get the user's Distinguished Name (DN)
$userDN = $user.DistinguishedName

# Get all groups the user is a member of (both security and distribution)
$allGroups = $user.MemberOf

# Initialize an array to hold report data
$report = @()

# Loop through all groups to determine their type
foreach ($groupDN in $allGroups) {
    $group = Get-ADGroup -Identity $groupDN -Property GroupScope, GroupCategory

    # Add an entry for each group (security or distribution) to the report array
    $report += [PSCustomObject]@{
        UserName          = "$firstName $lastName"
        DistinguishedName = $userDN
        SecurityGroups    = $group.Name
        GroupType         = $group.GroupCategory  # Identifies if it's a Security or Distribution group
    }
}

# Export results to a CSV file
$report | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "nReport saved to: $csvPath" -ForegroundColor Cyan