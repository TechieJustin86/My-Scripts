<#
This script is designed to extract Active Directory (AD) user information for all users within a specified Organizational Unit (OU) 
and export the data to an Excel file, with each user’s information placed on a separate worksheet within the Excel file.
The OU Distinguished Name (DN) will be pulled from the "UserAccountCreation.xlsx" file, from the "Domain" sheet under the "EnabledOU" column.
#>

# Define the path to the input Excel file and output file
$inputFile = "C:\Temp\UserAccountCreation.xlsx"
$outputFile = "C:\Temp\ADUsersExport.xlsx"

# Import the "Domain" sheet from the Excel file
$sheetNameDomainData = Import-Excel -Path $inputFile -WorksheetName "Domain"

# Extract the OU Distinguished Name from the "EnabledOU" column
$ouDistinguishedName = $sheetNameDomainData.EnabledOU

if (-not $ouDistinguishedName) {
    Write-Host "EnabledOU is not found in the Excel file or is empty."
    exit
}

# Specify the attributes to retrieve for each user
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

# Get all users in the specified OU
$users = Get-ADUser -Filter * -SearchBase $ouDistinguishedName -Property $attributes

if ($users.Count -eq 0) {
    Write-Host "No users found in the specified OU: $ouDistinguishedName"
    exit
}

# Create a new Excel file or remove the old one if it exists
if (Test-Path $outputFile) {
    Remove-Item $outputFile
}

Write-Host "Exporting users to $outputFile..."

# Export first user details to set up the header
$firstUser = $users[0]

# Resolve manager name (with error handling)
$managerName = $null
try {
    if ($firstUser.Manager) {
        $managerName = (Get-ADUser -Identity $firstUser.Manager).Name
    }
} catch {
    Write-Host "Could not resolve manager for $($firstUser.SamAccountName)"
}

# Prepare the first user's general information
$userDetails = [PSCustomObject]@{
    GivenName       = $firstUser.GivenName
    Surname         = $firstUser.Surname
    EmailAddress    = $firstUser.EmailAddress
    SamAccountName  = $firstUser.SamAccountName
    MobilePhone     = $firstUser.MobilePhone
    IPPhone         = $firstUser.IPPhone
    Title           = $firstUser.Title
    Department      = $firstUser.Department
    Manager         = $managerName
    MemberOf        = ""
}

# Prepare the group memberships for the first user
$groupMemberships = if ($firstUser.MemberOf) {
    $firstUser.MemberOf | ForEach-Object {
        $groupName = (Get-ADGroup -Identity $_).Name
        [PSCustomObject]@{
            GivenName       = ""
            Surname         = ""
            EmailAddress    = ""
            SamAccountName  = ""
            MobilePhone     = ""
            IPPhone         = ""
            Title           = ""
            Department      = ""
            Manager         = ""
            MemberOf        = $groupName
        }
    }
} else {
    @(
        [PSCustomObject]@{
            GivenName       = ""
            Surname         = ""
            EmailAddress    = ""
            SamAccountName  = ""
            MobilePhone     = ""
            IPPhone         = ""
            Title           = ""
            Department      = ""
            Manager         = ""
            MemberOf        = "No Group Memberships"
        }
    )
}

# Combine user details and group memberships
$sheetData = @($userDetails) + $groupMemberships

# Ensure a unique sheet name for the first user
$sheetName = "$($firstUser.SamAccountName)_$($firstUser.GivenName)"
$sheetData | Export-Excel -Path $outputFile -WorksheetName $sheetName -AutoSize -AutoFilter -Append

# Export remaining users without headers
foreach ($user in $users[1..$users.Count]) {
    # Resolve manager name (with error handling)
    $managerName = $null
    try {
        if ($user.Manager) {
            $managerName = (Get-ADUser -Identity $user.Manager).Name
        }
    } catch {
        Write-Host "Could not resolve manager for $($user.SamAccountName)"
    }

    # Prepare the user's general information
    $userDetails = [PSCustomObject]@{
        GivenName       = $user.GivenName
        Surname         = $user.Surname
        EmailAddress    = $user.EmailAddress
        SamAccountName  = $user.SamAccountName
        MobilePhone     = $user.MobilePhone
        IPPhone         = $user.IPPhone
        Title           = $user.Title
        Department      = $user.Department
        Manager         = $managerName
        MemberOf        = ""
    }

    # Prepare the group memberships for the user
    $groupMemberships = if ($user.MemberOf) {
        $user.MemberOf | ForEach-Object {
            $groupName = (Get-ADGroup -Identity $_).Name
            [PSCustomObject]@{
                GivenName       = ""
                Surname         = ""
                EmailAddress    = ""
                SamAccountName  = ""
                MobilePhone     = ""
                IPPhone         = ""
                Title           = ""
                Department      = ""
                Manager         = ""
                MemberOf        = $groupName
            }
        }
    } else {
        @(
            [PSCustomObject]@{
                GivenName       = ""
                Surname         = ""
                EmailAddress    = ""
                SamAccountName  = ""
                MobilePhone     = ""
                IPPhone         = ""
                Title           = ""
                Department      = ""
                Manager         = ""
                MemberOf        = "No Group Memberships"
            }
        )
    }

    # Combine user details and group memberships
    $sheetData = @($userDetails) + $groupMemberships

    # Ensure a unique sheet name for each user
    $sheetName = "$($user.SamAccountName)_$($user.GivenName)"
    $sheetData | Export-Excel -Path $outputFile -WorksheetName $sheetName -AutoSize -AutoFilter -Append
}

Write-Host "Export completed successfully. File saved at $outputFile"
