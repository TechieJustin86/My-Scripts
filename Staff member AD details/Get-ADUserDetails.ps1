<#
This script is designed to extract Active Directory (AD) user information for a given user 
and export the data to an Excel file.
#>


# Prompt the user for their name
$nameInput = Read-Host "Enter the First and Last Name of the user"

# Check if the input is empty
if (-not $nameInput) {
    Write-Host "You must enter both first and last name."
    exit
}

# Split the name input into first and last name parts
$nameParts = $nameInput -split '\s+' # Splits by any whitespace
if ($nameParts.Count -ne 2) {
    Write-Host "Please provide both first and last names separated by a single space."
    exit
}

$firstName = $nameParts[0]
$lastName = $nameParts[1]

# Define output file path
$outputFile = "C:\Temp\$($firstName)_$($lastName)_UserInfo.csv"

# Ensure the output directory exists
if (-not (Test-Path "C:\Temp")) {
    Write-Host "Directory 'C:\Temp' does not exist. Creating it now."
    New-Item -ItemType Directory -Force -Path "C:\Temp"
}

# Define the attributes to retrieve
$attributes = @(
    'GivenName',
    'Surname',
    'EmailAddress',
    'SamAccountName',
    'MobilePhone',
    'IPPhone',
    'Title',
    'Department',
    'Manager',
    'MemberOf'
)

# Function to add user data to the output array
function Add-UserData {
    param(
        [string]$givenName,
        [string]$surname,
        [string]$emailAddress,
        [string]$samAccountName,
        [string]$mobilePhone,
        [string]$ipPhone,
        [string]$title,
        [string]$department,
        [string]$manager,
        [string]$memberOf
    )

    $outputData += [PSCustomObject]@{
        GivenName      = $givenName
        Surname        = $surname
        EmailAddress   = $emailAddress
        SamAccountName = $samAccountName
        MobilePhone    = $mobilePhone
        IPPhone        = $ipPhone
        Title          = $title
        Department     = $department
        Manager        = $manager
        MemberOf       = $memberOf
    }
}

# Try to retrieve the user from Active Directory
try {
    $user = Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Property $attributes
} catch {
    Write-Host "An error occurred while querying Active Directory: $_"
    exit
}

# If user is found
if ($user) {

    # Retrieve manager details
    $managerName = if ($user.Manager) {
        (Get-ADUser -Identity $user.Manager).Name
    } else {
        "No Manager Assigned"
    }

    $outputData = @()

    # Add main user data
    Add-UserData -givenName $user.GivenName -surname $user.Surname -emailAddress $user.EmailAddress -samAccountName $user.SamAccountName -mobilePhone $user.MobilePhone -ipPhone $user.IPPhone -title $user.Title -department $user.Department -manager $managerName -memberOf "General Information"

    # If user belongs to groups, add each group to the output
    if ($user.MemberOf) {
        foreach ($group in $user.MemberOf) {
            $outputData += [PSCustomObject]@{
                GivenName      = ""
                Surname        = ""
                EmailAddress   = ""
                SamAccountName = ""
                MobilePhone    = ""
                IPPhone        = ""
                Title          = ""
                Department     = ""
                Manager        = ""
                MemberOf       = (Get-ADGroup -Identity $group).Name
            }
        }
    } else {
        # If no group memberships, log this
        Add-UserData -givenName "" -surname "" -emailAddress "" -samAccountName "" -mobilePhone "" -ipPhone "" -title "" -department "" -manager "" -memberOf "No Group Memberships"
    }

    # Export to CSV
    try {
        $outputData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        Write-Host "User information exported to $outputFile"
    } catch {
        Write-Host "An error occurred while exporting to CSV: $_"
    }

} else {
    Write-Host "User with First Name '$firstName' and Last Name '$lastName' not found in Active Directory."
}