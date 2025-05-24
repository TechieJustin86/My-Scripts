<#
This script is designed to extract Active Directory (AD) user information for all users within a specified Organizational Unit (OU) 
and export the data to an Excel file, with each user’s information placed on a separate worksheet within the Excel file
#>

# Hardcoded OU Distinguished Name
$ouDistinguishedName = "OU=XXX,DC=XXX,DC=XXX"

# Define the path for the output Excel file
$outputFile = "C:\folder\ADUsersExport.xlsx"

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
# Resolve manager name (if it exists) for the first user
$managerName = if ($firstUser.Manager) {
    (Get-ADUser -Identity $firstUser.Manager).Name
} else {
    $null
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

$sheetData = @($userDetails) + $groupMemberships
$sheetData | Export-Excel -Path $outputFile -WorksheetName $firstUser.SamAccountName -AutoSize -AutoFilter -Append

# Export remaining users without headers
foreach ($user in $users[1..$users.Count]) {
    $managerName = if ($user.Manager) {
        (Get-ADUser -Identity $user.Manager).Name
    } else {
        $null
    }

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

    # Combine user details and group memberships (fix applied)
    $sheetData = @($userDetails) + $groupMemberships

    # Export the data without headers
    $sheetData | Export-Excel -Path $outputFile -WorksheetName $user.SamAccountName -AutoSize -AutoFilter -Append
}

Write-Host "Export completed successfully. File saved at $outputFile"