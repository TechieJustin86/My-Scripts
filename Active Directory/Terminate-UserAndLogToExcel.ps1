# Define Excel file paths
$excelFile1 = "C:\Temp\TermainatedStaff.xlsx"
$excelFile2 = "C:\Temp\UserAccountCreation.xlsx"

# Load Excel workbooks
$excel1 = Open-ExcelPackage -Path $excelFile1
$worksheet1 = $excel1.Workbook.Worksheets["Terminated"]

$excel2 = Open-ExcelPackage -Path $excelFile2
$worksheet2 = $excel2.Workbook.Worksheets["Domain"]

# Ensure worksheets are loaded
if ($null -eq $worksheet1 -or $null -eq $worksheet2) {
    Write-Host "One or both worksheets could not be found. Please check worksheet names."
    Close-ExcelPackage $excel1
    Close-ExcelPackage $excel2
    exit
}

# Parse OU data from the "Domain" workbook
$UserOU = @{}
$sheetNameDomainData = Import-Excel -Path $excelFile2 -WorksheetName "Domain"

if ($sheetNameDomainData) {
    foreach ($row in $sheetNameDomainData) {
        if ($row.EnabledOU) {
            $UserOU["EnabledOU"] = $row.EnabledOU
        }
        if ($row.DisabledOU) {
            $UserOU["DisabledOU"] = $row.DisabledOU
        }
    }
}

# Assign OU values
$EnabledOU = $UserOU["EnabledOU"]
$DisabledOU = $UserOU["DisabledOU"]

# Function to generate a random password
function Generate-RandomPassword {
    $uppercase = -join ((65..90) | Get-Random -Count 5 | ForEach-Object {[char]$_})
    $lowercase = -join ((97..122) | Get-Random -Count 5 | ForEach-Object {[char]$_})
    $numbers = -join ((0..9) | Get-Random -Count 3)
    $symbols = -join ((33..47 + 58..64) | Get-Random -Count 2 | ForEach-Object {[char]$_})
    $password = $uppercase + $lowercase + $numbers + $symbols
    return -join ($password.ToCharArray() | Get-Random -Count $password.Length)
}

# Function to calculate a date 3 months from today
function Get-DateThreeMonthsFromNow {
    return (Get-Date).AddMonths(3).ToString("MM/dd/yyyy")
}


$firstAndLastName = Read-Host "Enter the First and Last Name of the terminated employee (First Last)"
$nameParts = $firstAndLastName -split '\s+'
    if ($nameParts.Count -ne 2) {
    Write-Host "Invalid input! Please provide both first and last names separated by a space."
    }

$firstName = $nameParts[0]
$lastName = $nameParts[1]

# Input Delegate Email (Optional)
$delegateEmail = Read-Host "Enter the delegate email (leave blank if not applicable)"

# Input option to keep the account enabled (Y/N)
$enableAccount = (Read-Host "Keep Account Enabled (Y or N)").Trim().ToUpper()
if ($enableAccount -ne 'Y' -and $enableAccount -ne 'N') {
    Write-Host "Invalid input! Please enter 'Y' or 'N'."
    exit
}
$enableAccount = $enableAccount -eq 'Y'

# Input option to keep group memberships (Y/N)
$keepGroups = (Read-Host "Keep Group Memberships (Y or N)").Trim().ToUpper()
if ($keepGroups -ne 'Y' -and $keepGroups -ne 'N') {
    Write-Host "Invalid input! Please enter 'Y' or 'N'."
    exit
}
$keepGroups = $keepGroups -eq 'Y'

Write-Host "Searching for user: First Name: $firstName, Last Name: $lastName"

do {
    $account = Get-ADUser -Filter {GivenName -eq $firstName -and Surname -eq $lastName} -SearchBase $EnabledOU

    if (!$account) {
        Write-Host "User not found in the specified OU: $EnabledOU. Please check the name and try again."
        $retry = Read-Host "Do you want to retry? (Y/N)"
        if ($retry -ne 'Y') { exit }
        $firstAndLastName = Read-Host "Enter the First and Last Name of the terminated employee (First Last)"
        $nameParts = $firstAndLastName -split '\s+'
        $firstName = $nameParts[0]
        $lastName = $nameParts[1]
    }
} while (!$account)


# Generate random password and update the account
$password = Generate-RandomPassword
Set-ADAccountPassword -Identity $account.SamAccountName -Reset -NewPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Confirm:$false

# Hide from address list
Set-ADUser -Identity $account.SamAccountName -Replace @{msExchHideFromAddressLists=$true} -Confirm:$false

# Add today's date to Notes in Telephones tab (using -Replace to update the 'Info' field)
$today = Get-Date -Format "MM/dd/yyyy"
Set-ADUser -Identity $account.SamAccountName -Replace @{Info="Terminated on $today"} -Confirm:$false

$UserSid = [PSCustomObject]@{
    Name              = $account.Name
    UserPrincipalName = $account.UserPrincipalName
    SAMAccountName    = $account.SamAccountName
    SID               = $account.SID.Value
}

# Remove all groups except 'Domain Users' if checkbox for removing groups is unchecked
if (-not $keepGroups) {
    $groups = Get-ADUser $account.SamAccountName -Property MemberOf | Select-Object -ExpandProperty MemberOf
    foreach ($group in $groups) {
        if ($group -notlike '*Domain Users*') {
            Remove-ADGroupMember -Identity $group -Members $account.SamAccountName -Confirm:$false
        }
    }
}

# Check if the account should be enabled or disabled based on the input
if (-not $enableAccount) {
    # Disable account
    Disable-ADAccount -Identity $account.SamAccountName -Confirm:$false

    # Move the account to Disabled Employees OU
    Move-ADObject -Identity $account.DistinguishedName -TargetPath $DisabledOU -Confirm:$false

    # Remove specific attribute values if account is disabled
    $attributesToClear = @("company", "department", "description", "facsimileTelephoneNumber", "mail", "manager", "streetAddress", "telephoneNumber", "title", "wWWHomePage")

    foreach ($attr in $attributesToClear) {
        $value = (Get-ADUser $account.SamAccountName -Properties $attr).$attr
        if ($value -ne $null -and $value -ne "") {
            Set-ADUser -Identity $account.SamAccountName -Clear $attr -Confirm:$false
        }
    }

    # Set extensionAttribute5 to "Terminated - No Email"
    Set-ADUser -Identity $account.SamAccountName -Replace @{extensionAttribute5="Terminated - No Email"} -Confirm:$false

    Write-Host "Account has been disabled and moved to Disabled Employees OU."
} else {
    # Enable account
    Enable-ADAccount -Identity $account.SamAccountName -Confirm:$false

    # Calculate the date 3 months from today
    $removeEmailAccessDate = Get-DateThreeMonthsFromNow

    # Set the description to "Remove email access on <3 months from today>"
    Set-ADUser -Identity $account.SamAccountName -Replace @{description="Remove email access on $removeEmailAccessDate"} -Confirm:$false

    # Set extensionAttribute5 to "Terminated - Email"
    Set-ADUser -Identity $account.SamAccountName -Replace @{extensionAttribute5="Terminated - Email"} -Confirm:$false

    # Only add Mailbox Permission if a delegate is specified
    if (![string]::IsNullOrEmpty($delegateEmail)) {
        Add-MailboxPermission -Identity $account.SamAccountName -User $delegateEmail -AccessRights FullAccess -InheritanceType All
        Write-Host "Email Delegation to: $delegateEmail"
    } else {
        Write-Host "No delegate specified, skipping email delegation."
    }

    Write-Host "Account has been enabled and remains in the current OU."
}

# Get the next available row number in the "Terminated" worksheet
$startRow = $worksheet1.Dimension.End.Row + 1

# Write termination details to the "Terminated" worksheet
$worksheet1.Cells[$startRow, 1].Value = "$firstName $lastName"
$worksheet1.Cells[$startRow, 2].Value = $account.SamAccountName
$worksheet1.Cells[$startRow, 3].Value = if ($enableAccount) { "Enabled" } else { "Disabled" }
$worksheet1.Cells[$startRow, 4].Value = if ($keepGroups) { "Kept" } else { "Removed" }
$worksheet1.Cells[$startRow, 5].Value = "Yes"
$worksheet1.Cells[$startRow, 6].Value = if ($enableAccount) { "Email-Enabled" } else { "Email-Disabled" }
$worksheet1.Cells[$startRow, 7].Value = "Terminated on $today"
$worksheet1.Cells[$startRow, 8].Value = if ($enableAccount) { "Remove email access on $removeEmailAccessDate" } else { "Not applicable (account disabled)" }
$worksheet1.Cells[$startRow, 9].Value = $password
$worksheet1.Cells[$startRow, 10].Value = if (![string]::IsNullOrEmpty($delegateEmail)) { "Delegated to: $delegateEmail" } else { "No delegation" }
$worksheet1.Cells[$startRow, 11].Value = $UserSid.SID


# Save and close Excel files
try {
    $excel1.Save()
    $excel2.Save()
    Close-ExcelPackage $excel1
    Close-ExcelPackage $excel2
    Write-Host "Data successfully exported to Excel."
} catch {
    Write-Host "Failed to save or close the Excel file. Error: $_"
    exit
}
