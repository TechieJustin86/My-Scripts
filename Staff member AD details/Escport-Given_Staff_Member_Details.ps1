<#
This script is designed to extract Active Directory (AD) user information for a given user 
and export the data to an Excel file.
#>

$nameInput = Read-Host "Enter the First and Last Name of the user"

$nameParts = $nameInput -split '\s+' # Splits by any whitespace
if ($nameParts.Count -ne 2) {
    Write-Host "Please provide both first and last names separated by a space!"
    exit
}

$firstName = $nameParts[0]
$lastName = $nameParts[1]

$outputFile = "C:\Temp\$($user.SamAccountName).csv"

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

$user = Get-ADUser -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName'" -Property $attributes

if ($user) {

    $managerName = if ($user.Manager) {
        (Get-ADUser -Identity $user.Manager).Name
    } else {
        $null
    }

    $outputData = @()

    $outputData += [PSCustomObject]@{
        GivenName       = $user.GivenName
        Surname         = $user.Surname
        EmailAddress    = $user.EmailAddress
        SamAccountName  = $user.SamAccountName
        MobilePhone     = $user.MobilePhone
        IPPhone         = $user.IPPhone
        Title           = $user.Title
        Department      = $user.Department
        Manager         = $managerName
        MemberOf        = "General Information"
    }

    if ($user.MemberOf) {
        foreach ($group in $user.MemberOf) {
            $outputData += [PSCustomObject]@{
                GivenName       = ""
                Surname         = ""
                EmailAddress    = ""
                SamAccountName  = ""
                MobilePhone     = ""
                IPPhone         = ""
                Title           = ""
                Department      = ""
                Manager         = ""
                MemberOf        = (Get-ADGroup -Identity $group).Name
            }
        }
    } else {
        $outputData += [PSCustomObject]@{
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
    }

    # Export to CSV
    $outputData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
    Write-Host "User information exported to $outputFile"
} else {
    Write-Host "User with First Name '$firstName' and Last Name '$lastName' not found in Active Directory."
}
