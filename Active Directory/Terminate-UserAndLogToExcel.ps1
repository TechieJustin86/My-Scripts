<#
This PowerShell script manages the termination process for an employee in Active Directory and updates the status in an Excel file. 
It checks and installs required modules, handles account management (disabling/enabling, password reset, group membership removal), 
updates Exchange Online mailbox permissions, and records the employee's termination details in the specified Excel sheet.

You MUST have the Sharepoint folder Synced to your computer or point $excelFile to the file location.

#>
# Check if the required module are available
$modules = @("ImportExcel", "ActiveDirectory", "ExchangeOnlineManagement")

foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "$module module is not installed. Installing now."
        try {
            Install-Module -Name $module -Force -Scope CurrentUser
        } catch {
            Write-Host "Failed to install $module module. Error: $($_.Exception.Message)"
            exit
        }
    }
}

# Set Execution Policy to bypass for the current session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Try to connect to Exchange Online
try {
    Connect-ExchangeOnline
} catch {
    Write-Host "Failed to connect to Exchange Online. Error: $_"
    exit
}

# Get logged-in user's name for Sharpoint mapping. Comment out if you not going to use sharepoint.
$loggedInUser = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName).Split("\")[-1]

# Path to the Excel file. Edit this to the file path if needed. The file is TermainatedStaff
$excelFile = "C:\Users\$loggedInUser\TermainatedStaff.xlsx"
$worksheetName = "Terminated"

# Enabled OU
$EnabledOU = "OU=XXX,DC=XXX,DC=XXX"

# Disabled OU
$DisabledOU = "OU=XXX,DC=XXX,DC=XXX"

# Open the Excel file
$excel = Open-ExcelPackage -Path $excelFile
$worksheet = $excel.Workbook.Worksheets[$worksheetName]

# Check if the worksheet exists
if ($null -eq $worksheet) {
    Write-Host "Worksheet '$worksheetName' not found in the Excel file."
    exit
}

# Get the next available row number to append new data
$startRow = $worksheet.Dimension.End.Row + 1

# Function to generate random password (15 characters: 5 uppercase, 5 lowercase, 3 numbers, 2 symbols)
function Generate-RandomPassword {
    $uppercase = -join ((65..90) | Get-Random -Count 5 | ForEach-Object {[char]$_})
    $lowercase = -join ((97..122) | Get-Random -Count 5 | ForEach-Object {[char]$_})
    $numbers = -join ((0..9) | Get-Random -Count 3)
    $symbols = -join ((33..47 + 58..64) | Get-Random -Count 2 | ForEach-Object {[char]$_})
    $password = $uppercase + $lowercase + $numbers + $symbols
    return -join ($password.ToCharArray() | Get-Random -Count $password.Length)
}

# Function to calculate date 3 months from today
function Get-DateThreeMonthsFromNow {
    return (Get-Date).AddMonths(3).ToString("MM/dd/yyyy")
}

# Input First and Last Name of the terminated employee
$firstAndLastName = Read-Host "Enter the First and Last Name of the terminated employee (First Last)"

$nameParts = $firstAndLastName -split '\s+' # Splits by any whitespace
if ($nameParts.Count -ne 2) {
    Write-Host "Please provide both first and last names separated by a space!"
    exit
}

$firstName = $nameParts[0]
$lastName = $nameParts[1]

# Input Delegate Email (Optional)
$delegateEmail = Read-Host "Enter the delegate email (leave blank if not applicable)"

# Input option to keep the account enabled (Y/N)
$enableAccount = Read-Host "Keep Account Enabled (Y or N)"
if ($enableAccount -ne 'Y' -and $enableAccount -ne 'N') {
    Write-Host "Invalid input! Please enter 'Y' or 'N'."
    exit
}
$enableAccount = $enableAccount -eq 'Y'

# Input option to keep group memberships (Y/N)
$keepGroups = Read-Host "Keep Group Memberships (Y or N)"
if ($keepGroups -ne 'Y' -and $keepGroups -ne 'N') {
    Write-Host "Invalid input! Please enter 'Y' or 'N'."
    exit
}
$keepGroups = $keepGroups -eq 'Y'

# Get user account
$account = Get-ADUser -Filter {GivenName -eq $firstName -and Surname -eq $lastName} -SearchBase $EnabledOU

if (!$account) {
    Write-Host "User not found in the specified OU!"
    exit
}

# Generate random password and set it
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

# Write data to Excel
$worksheet.Cells[$startRow, 1].Value = "$firstName $lastName"
$worksheet.Cells[$startRow, 2].Value = $account.SamAccountName
$worksheet.Cells[$startRow, 3].Value = if ($enableAccount) { "Enabled" } else { "Disabled" }
$worksheet.Cells[$startRow, 4].Value = if ($keepGroups) { "Kept" } else { "Removed" }
$worksheet.Cells[$startRow, 5].Value = "Yes"
$worksheet.Cells[$startRow, 6].Value = if ($enableAccount) { "Email-Enabled" } else { "Email-Disabled" }
$worksheet.Cells[$startRow, 7].Value = "Terminated on $today"
$worksheet.Cells[$startRow, 8].Value = if ($enableAccount) { "Remove email access on $removeEmailAccessDate" } else { "Not applicable (account disabled)" }
$worksheet.Cells[$startRow, 9].Value = $password
$worksheet.Cells[$startRow, 10].Value = if (![string]::IsNullOrEmpty($delegateEmail)) { "Delegated to: $delegateEmail" } else { "No delegation" }
$worksheet.Cells[$startRow, 11].Value = $UserSid.SID

try {
    $excel.Save()
    Close-ExcelPackage $excel
    Write-Host "Data successfully exported to Excel."
} catch {
    Write-Host "Failed to save or close the Excel file. Error: $_"
    exit
}
