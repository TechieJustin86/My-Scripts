<#
This PowerShell script manages the termination process for an employee in Active Directory and updates the status in an Excel file. 
It checks and installs required modules, handles account management (disabling/enabling, password reset, group membership removal), 
updates Exchange Online mailbox permissions, and records the employee's termination details in the specified Excel sheet.

You can use SharePoint folder to store the file or point $excelFile to the file location.

#>

# Check if the required modules are available and import them
$modules = @("ImportExcel", "ActiveDirectory", "ExchangeOnlineManagement")

foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "$module module is not installed. Installing now..."
        try {
            Install-Module -Name $module -Force -Scope CurrentUser
            Write-Host "$module installed successfully."
        } catch {
            Write-Warning "Failed to install $module module. Please install it manually. Error: $($_.Exception.Message)"
            exit
        }
    }
    Import-Module $module -ErrorAction Stop
}

# Set Execution Policy to bypass for the current session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Connect to Exchange Online with error handling
Write-Host "Connecting to Exchange Online..."
try {
    Connect-ExchangeOnline -ErrorAction Stop
    Write-Host "Connected to Exchange Online."
} catch {
    Write-Warning "Failed to connect to Exchange Online. Error: $_"
    exit
}

# Get logged-in user's name for SharePoint mapping. Comment out if you not going to use SharePoint.
$loggedInUser = $env:USERNAME

# Path to the Excel file. Edit this to the file path if needed.
$terminatedStaffFilePath  = "C:\Users\$loggedInUser\Staff-Termainated.xlsx"
$domainOUsFilePath  = "C:\Temp\Staff-AccountData.xlsx"
$worksheetName = "Terminated"
$sheetNameDomain = "Domain"

# Import domain OUs data from Excel
try {
    $sheetNameDomainData = Import-Excel -Path $domainOUsFilePath  -WorksheetName $sheetNameDomain -ErrorAction Stop
} catch {
    Write-Host "Failed to import domain sheet data from '$domainOUsFilePath'. Error: $_"
    exit
}

# Parse domain data into a hash table
$TerminateOU = @{
    EnabledOU = $null
    DisabledOU = $null
}

if ($sheetNameDomainData) {
    foreach ($row in $sheetNameDomainData) {
        if ($row.EnabledOU) {
            $TerminateOU["EnabledOU"] = $row.EnabledOU
        }
        if ($row.DisabledOU) {
            $TerminateOU["DisabledOU"] = $row.DisabledOU
        }
    }
}

# Assign OU variables
$EnabledOU = $TerminateOU["EnabledOU"]
$DisabledOU = $TerminateOU["DisabledOU"]

if (-not $EnabledOU -or -not $DisabledOU) {
    Write-Host "EnabledOU or DisabledOU not found in the domain sheet."
    exit
}

# Open Termainated Staff file for updating.
try {
    $excel = Open-ExcelPackage -Path $terminatedStaffFilePath -ErrorAction Stop
    $worksheet = $excel.Workbook.Worksheets[$worksheetName]
} catch {
    Write-Host "Failed to open the Excel file '$terminatedStaffFilePath'. Error: $_"
    exit
}

if (-not $worksheet) {
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
$firstAndLastName = Read-Host "Enter the First and Last Name of the terminated employee."

$nameParts = $firstAndLastName -split '\s+' # Splits by any whitespace
if ($nameParts.Count -ne 2) {
    Write-Host "Please provide both first and last names separated by a space!"
    exit
}

$firstName = $nameParts[0]
$lastName = $nameParts[1]

# Input Delegate Email if applicable.
$delegateEmail = Read-Host "Enter the delegate email (leave blank if not applicable)"

# Input option to keep the account enabled (Y/N)
$enableAccountInput = Read-Host "Keep Account Enabled (Y or N)"
if ($enableAccountInput -ne 'Y' -and $enableAccountInput -ne 'N') {
    Write-Host "Invalid input! Please enter 'Y' or 'N'."
    exit
}
$enableAccount = $enableAccountInput -eq 'Y'

# Input option to keep group memberships (Y/N)
$keepGroupsInput = Read-Host "Keep Group Memberships (Y or N)"
if ($keepGroupsInput -ne 'Y' -and $keepGroupsInput -ne 'N') {
    Write-Host "Invalid input! Please enter 'Y' or 'N'."
    exit
}
$keepGroups = $keepGroupsInput -eq 'Y'

# Get user account with error handling
try {
    $account = Get-ADUser -Filter {GivenName -eq $firstName -and Surname -eq $lastName} -SearchBase $EnabledOU -ErrorAction Stop
} catch {
    Write-Host "User not found in the specified OU! Error: $_"
    exit
}

if (-not $account) {
    Write-Host "User not found in the specified OU!"
    exit
}

# Reset password with generated created one.
$password = Generate-RandomPassword
try {
    Set-ADAccountPassword -Identity $account.SamAccountName -Reset -NewPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Confirm:$false -ErrorAction Stop
} catch {
    Write-Host "Failed to reset password. Error: $_"
    exit
}

# Hide from Office 365 address list.
try {
    Set-ADUser -Identity $account.SamAccountName -Replace @{msExchHideFromAddressLists=$true} -Confirm:$false -ErrorAction Stop
} catch {
    Write-Host "Failed to hide user from address lists. Error: $_"
}

# Add today's date to Notes in Telephones tab (using -Replace to update the 'Info' field)
$today = Get-Date -Format "MM/dd/yyyy"
try {
    Set-ADUser -Identity $account.SamAccountName -Replace @{Info="Terminated on $today"} -Confirm:$false -ErrorAction Stop
} catch {
    Write-Host "Failed to update Info field. Error: $_"
}

$UserSid = [PSCustomObject]@{
    Name              = $account.Name
    UserPrincipalName = $account.UserPrincipalName
    SAMAccountName    = $account.SamAccountName
    SID               = $account.SID.Value
}

# Remove all groups except 'Domain Users' if checkbox for removing groups is unchecked
if (-not $keepGroups) {
    try {
        $groups = Get-ADUser $account.SamAccountName -Property MemberOf -ErrorAction Stop | Select-Object -ExpandProperty MemberOf
        foreach ($group in $groups) {
            $groupName = (Get-ADGroup -Identity $group).Name
            if ($groupName -ne 'Domain Users') {
                Remove-ADGroupMember -Identity $group -Members $account.SamAccountName -Confirm:$false -ErrorAction Stop
            }
        }
    } catch {
        Write-Host "Failed to remove group memberships. Error: $_"
    }
}

# Check if the account should be enabled or disabled based on the input
if (-not $enableAccount) {
    # Disable account
    try {
        Disable-ADAccount -Identity $account.SamAccountName -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to disable account. Error: $_"
    }

    # Move the account to Disabled Employees OU
    try {
        Move-ADObject -Identity $account.DistinguishedName -TargetPath $DisabledOU -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to move account to Disabled OU. Error: $_"
    }

    # Remove specific attribute values if account is disabled
    $attributesToClear = @("company", "department", "description", "facsimileTelephoneNumber", "mail", "manager", "streetAddress", "telephoneNumber", "title", "wWWHomePage")
    foreach ($attr in $attributesToClear) {
        try {
            $value = (Get-ADUser $account.SamAccountName -Properties $attr -ErrorAction Stop).$attr
            if ($value -ne $null -and $value -ne "") {
                Set-ADUser -Identity $account.SamAccountName -Clear $attr -Confirm:$false -ErrorAction Stop
            }
        } catch {
            Write-Host "Failed to clear attribute '$attr'. Error: $_"
        }
    }

    # Set extensionAttribute5 to "Terminated - No Email"
    try {
        Set-ADUser -Identity $account.SamAccountName -Replace @{extensionAttribute5="Terminated - No Email"} -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to set extensionAttribute5. Error: $_"
    }

    Write-Host "Account has been disabled and moved to Disabled Employees OU."
} else {
    # Enable account
    try {
        Enable-ADAccount -Identity $account.SamAccountName -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to enable account. Error: $_"
    }

    # Calculate the date 3 months from today
    $removeEmailAccessDate = Get-DateThreeMonthsFromNow

    # Set the description to "Remove email access on <3 months from today>"
    try {
        Set-ADUser -Identity $account.SamAccountName -Replace @{description="Remove email access on $removeEmailAccessDate"} -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to update description. Error: $_"
    }

    # Set extensionAttribute5 to "Terminated - Email"
    try {
        Set-ADUser -Identity $account.SamAccountName -Replace @{extensionAttribute5="Terminated - Email"} -Confirm:$false -ErrorAction Stop
    } catch {
        Write-Host "Failed to set extensionAttribute5. Error: $_"
    }

    # Only add Mailbox Permission if a delegate is specified
    if (![string]::IsNullOrEmpty($delegateEmail)) {
        try {
            Add-MailboxPermission -Identity $account.SamAccountName -User $delegateEmail -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
            Write-Host "Email Delegation to: $delegateEmail"
        } catch {
            Write-Host "Failed to add mailbox permission. Error: $_"
        }
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