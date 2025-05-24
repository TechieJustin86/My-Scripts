<#
This script is useful for AD administrators who want to track and report the password age of users in a specified OU or domain. 
It outputs the results both in the console and optionally saves them to a CSV file for further analysis.
#>

# Import the Active Directory module
Import-Module ActiveDirectory

# Specify the Organizational Unit (OU) or leave it blank to search the entire domain
$OU = "OU=XXX,DC=XXX,DC=org"  # Replace with your OU or domain

# Get all users from the specified OU
try {
    $users = Get-ADUser -Filter * -SearchBase $OU -Property DisplayName, SamAccountName, pwdLastSet
} catch {
    Write-Host "Error retrieving users from Active Directory: $_" -ForegroundColor Red
    return
}

# Create a collection to store results
$results = @()

# Loop through each user and calculate the password age
foreach ($user in $users) {
    # Convert the pwdLastSet attribute to a readable date
    $pwdLastSet = [datetime]::FromFileTimeUtc($user.pwdLastSet)

    # Calculate password age in days
    $passwordAgeDays = (Get-Date) - $pwdLastSet
    $passwordAgeDays = $passwordAgeDays.Days

    # Add the result to the collection
    $results += [PSCustomObject]@{
        DisplayName      = $user.DisplayName
        SamAccountName   = $user.SamAccountName
        PasswordLastSet  = $pwdLastSet
        PasswordAgeDays  = $passwordAgeDays
    }
}

# Output results to console in a table format
$results | Format-Table -Property DisplayName, SamAccountName, PasswordLastSet, PasswordAgeDays -AutoSize

# Optionally export to a CSV file
$csvPath = "C:\Temp\PasswordAgeReport.csv"
try {
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Password age report has been saved to $csvPath" -ForegroundColor Green
} catch {
    Write-Host "Error saving the report: $_" -ForegroundColor Red
}