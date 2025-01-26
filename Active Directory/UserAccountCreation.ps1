# Path to Excel File
$excelFile = "C:\Temp\UserAccountCreation.xlsx"

# Sheets to load
$sheetNameGroups = "Job Title Groups"
$sheetNameUserDetails = "Department Manager"
$sheetNameDepartmentOU = "Department OU"
$sheetNameOfficeOU = "Office OU"
$sheetNameOfficeLocation = "Office Location"
$sheetNameDomain = "Domain"

# Load data from Excel sheets
$titleData = Import-Excel -Path $excelFile -WorksheetName $sheetNameGroups
$userDetails = Import-Excel -Path $excelFile -WorksheetName $sheetNameUserDetails
$departmentOUData = Import-Excel -Path $excelFile -WorksheetName $sheetNameDepartmentOU
$officeOUData = Import-Excel -Path $excelFile -WorksheetName $sheetNameOfficeOU
$officeLocationData = Import-Excel -Path $excelFile -WorksheetName $sheetNameOfficeLocation
$sheetNameDomainData = Import-Excel -Path $excelFile -WorksheetName $sheetNameDomain


# Extract titles from the "Groups" sheet (column headers)
$titles = $titleData | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name

# Get unique manager names from the "Manager" column in the "User Details" sheet
$managers = $userDetails | Select-Object -ExpandProperty Manager | Where-Object { $_ -ne $null } | Sort-Object -Unique

# Extract unique department names from the "Department" column in the "User Details" sheet
$departments = $userDetails | Select-Object -ExpandProperty Department | Where-Object { $_ -ne $null } | Sort-Object -Unique

#region parsing data into a hash table

# Parse DepartmentOU data into a hash table
$DepartmentOUs = @{}
if ($departmentOUData) {
    foreach ($row in $departmentOUData) {
        if ($row.Department -and $row.OUPath) {
            $DepartmentOUs[$row.Department] = $row.OUPath
        }
    }
}

# Parse OfficeOU data into a hash table
$OfficeOUs = @{}
if ($officeOUData) {
    foreach ($row in $officeOUData) {
        if ($row.Office -and $row.OUPath) {
            $OfficeOUs[$row.Office] = $row.OUPath
        }
    }
}

# Parse OfficeLocation data into a hash table
$OfficeLocations = @{}
if ($officeLocationData) {
    foreach ($row in $officeLocationData) {
        if ($row.Department -and $row.Office) {
            $OfficeLocations[$row.Department] = $row.Office
        }
    }
}

# Parse domain data into a hash table
$Domains = @{}
if ($sheetNameDomainData) {
    foreach ($row in $sheetNameDomainData) {
        if ($row.Domainorg) {
            $Domains["Domainorg"] = $row.Domainorg
        }
        if ($row.Domaincom) {
            $Domains["Domaincom"] = $row.Domaincom
        }
    }
}

# Parse company data into a hash table
$Companies = @{}
if ($sheetNameDomainData) {
    foreach ($row in $sheetNameDomainData) {
        if ($row.Company) {
            $Companies[$row.Company] = $row.Company
        }
    }
}

#endregion

# Array to collect output messages
$outputMessages = @()

Function Show-UserInputForm {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "User Information"
    $form.Size = New-Object System.Drawing.Size(300, 375)
    $form.StartPosition = 'CenterScreen'

    # Full Name Label and TextBox
    $FullNameLabel = New-Object System.Windows.Forms.Label
    $FullNameLabel.Text = "Full Name:"
    $FullNameLabel.AutoSize = $true
    $FullNameLabel.Top = 10
    $FullNameLabel.Left = 10
    $form.Controls.Add($FullNameLabel)

    $FullNameTextBox = New-Object System.Windows.Forms.TextBox
    $FullNameTextBox.Top = 40
    $FullNameTextBox.Left = 10
    $FullNameTextBox.Width = 100
    $form.Controls.Add($FullNameTextBox)

    # Security digits
    $SecuritydigitsLabel = New-Object System.Windows.Forms.Label
    $SecuritydigitsLabel.Text = "Enter 4 random digits:"
    $SecuritydigitsLabel.AutoSize = $true
    $SecuritydigitsLabel.Top = 10
    $SecuritydigitsLabel.Left = 120
    $form.Controls.Add($SecuritydigitsLabel)

    $SecuritydigitsBox = New-Object System.Windows.Forms.TextBox
    $SecuritydigitsBox.Top = 40
    $SecuritydigitsBox.Left = 120
    $SecuritydigitsBox.Width = 100
    $form.Controls.Add($SecuritydigitsBox)

    # Title Label and ComboBox
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Title:"
    $titleLabel.AutoSize = $true
    $titleLabel.Top = 70
    $titleLabel.Left = 10
    $form.Controls.Add($titleLabel)

    $titleComboBox = New-Object System.Windows.Forms.ComboBox
    $titleComboBox.Items.Add("")
    $titleComboBox.Items.AddRange($titles)
    $titleComboBox.SelectedIndex = 0
    $titleComboBox.Top = 90
    $titleComboBox.Left = 10
    $titleComboBox.Width = 200
    $form.Controls.Add($titleComboBox)

    # Manager Label and ComboBox
    $managerLabel = New-Object System.Windows.Forms.Label
    $managerLabel.Text = "Manager:"
    $managerLabel.AutoSize = $true
    $managerLabel.Top = 120
    $managerLabel.Left = 10
    $form.Controls.Add($managerLabel)

    $managerComboBox = New-Object System.Windows.Forms.ComboBox
    $managerComboBox.Items.Add("")
    $managerComboBox.Items.AddRange($managers)
    $managerComboBox.SelectedIndex = 0
    $managerComboBox.Top = 140
    $managerComboBox.Left = 10
    $managerComboBox.Width = 200
    $form.Controls.Add($managerComboBox)

    # Department Label and ComboBox
    $departmentLabel = New-Object System.Windows.Forms.Label
    $departmentLabel.Text = "Department:"
    $departmentLabel.AutoSize = $true
    $departmentLabel.Top = 170  # Moved down for spacing
    $departmentLabel.Left = 10
    $form.Controls.Add($departmentLabel)

    $departmentComboBox = New-Object System.Windows.Forms.ComboBox
    $departmentComboBox.Items.Add("")  # Blank option first
    $departmentComboBox.Items.AddRange($DepartmentOUs.Keys)
    $departmentComboBox.SelectedIndex = 0
    $departmentComboBox.Top = 190  # Moved down for spacing
    $departmentComboBox.Left = 10
    $departmentComboBox.Width = 200
    $form.Controls.Add($departmentComboBox)

    # Office Label and ComboBox
    $officeLabel = New-Object System.Windows.Forms.Label
    $officeLabel.Text = "Office:"
    $officeLabel.AutoSize = $true
    $officeLabel.Top = 220  # Moved down for spacing
    $officeLabel.Left = 10
    $form.Controls.Add($officeLabel)

    $officeComboBox = New-Object System.Windows.Forms.ComboBox
    $officeComboBox.Items.Add("")  # Blank option first
    $officeComboBox.Items.AddRange($officeOUs.Keys)
    $officeComboBox.SelectedIndex = 0
    $officeComboBox.Top = 240  # Moved down for spacing
    $officeComboBox.Left = 10
    $officeComboBox.Width = 100
    $form.Controls.Add($officeComboBox)

    # OK Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Top = 300 
    $okButton.Left = 10
    $okButton.Width = 80
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Top = 300  # Moved down for spacing
    $cancelButton.Left = 175
    $cancelButton.Width = 80
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Show the form and capture input
    $formResult = $form.ShowDialog()
    
    # If OK is clicked, return input; otherwise, return $null
    if ($formResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedOffice = $officeComboBox.SelectedItem
        $ouPath = if ($selectedOffice) { $officeOUs[$selectedOffice] } else { $null }
        return @{
            FullName = $FullNameTextBox.Text
            Securitydigits = $SecuritydigitsBox.Text
            Title = $titleComboBox.SelectedItem
            Manager = $managerComboBox.SelectedItem
            Department = $departmentComboBox.SelectedItem
            OUPath = $ouPath  # Include OUPath in the returned object
        }
    }
    
    return $null  # User cancelled
}

# Main script execution
$userInput = Show-UserInputForm

# Validate user input
if (-not $userInput.FullName) {
    Write-Host "Full Name is required."
    return
}

$nameParts = $userInput.FullName -split ' '
if ($nameParts.Length -ge 2) {
    $firstName = $nameParts[0]
    $lastName = $nameParts[1]
} else {
    $firstName = $nameParts[0]
    $lastName = ""
}

# Function to generate random password (15 characters: 5 uppercase, 5 lowercase, 3 numbers, 2 symbols)
function Generate-RandomPassword {
    $uppercase = -join ((65..90) | Get-Random -Count 5 | ForEach-Object {[char]$_})
    $lowercase = -join ((97..122) | Get-Random -Count 5 | ForEach-Object {[char]$_})
    $numbers = -join ((0..9) | Get-Random -Count 3)
    $symbols = -join ((33..47 + 58..64) | Get-Random -Count 2 | ForEach-Object {[char]$_})
    $password = $uppercase + $lowercase + $numbers + $symbols
    return -join ($password.ToCharArray() | Get-Random -Count $password.Length)
}
$password = Generate-RandomPassword


# Create the sAMAccountName and userPrincipalName
$securitydigits = $userInput.Securitydigits
$sAMAccountName = "$firstName$securitydigits"
$userPrincipalName = "$sAMAccountName@$($Domains["Domainorg"])"
$userEmail = "$firstName@$($Domains["Domaincom"])"
$Company = "$Companies"

# Use Domaincom for email
if ($email) {
    $selectedDomain = $Domains["Domaincom"]
    $userEmail = "$email@$selectedDomain"
}

# Use Domainorg for sAMAccountName
if ($sAMAccountName) {
    $selectedDomain = $Domains["Domainorg"]
    $userPrincipalName = "$sAMAccountName@$selectedDomain"
}



# Assign OU based on department or office selection
$selectedDepartment = $userInput.Department
$selectedOffice = $userInput.OUPath  # Assuming you are storing the selected office as OUPath

# Check if office is selected and map to department OU
if ($selectedOffice -and $OfficeLocations.ContainsValue($selectedOffice)) {
    $departmentFromOffice = $OfficeLocations.GetEnumerator() | Where-Object { $_.Value -eq $selectedOffice } | Select-Object -ExpandProperty Key
    
    # Now assign the OU path based on the department from the office
    if ($DepartmentOUs.ContainsKey($departmentFromOffice)) {
        $ouPath = $DepartmentOUs[$departmentFromOffice]
        $outputMessages += "Office $selectedOffice corresponds to department $departmentFromOffice with OU path: $ouPath"
    } else {
        $ouPath = $null  # Fallback if no matching department found
        $outputMessages += "No matching department found for office $selectedOffice."
    }
} elseif ($selectedDepartment -and $DepartmentOUs.ContainsKey($selectedDepartment)) {
    # Fallback to department if office isn't selected or doesn't map correctly
    $ouPath = $DepartmentOUs[$selectedDepartment]
    $outputMessages += "Department $selectedDepartment has OU path: $ouPath"
} else {
    # Default case if no valid department or office
    $ouPath = $null
    $outputMessages += "No matching department or office found for OU assignment."
}

# Search for the user in Active Directory by First and Last Name
$user = Get-ADUser -Filter { GivenName -eq $firstName -and Surname -eq $lastName }

# Define AD attributes #
$Title = $userInput.Title
$Department = $selectedDepartment
$managerName = $userInput.Manager
$manager = Get-ADUser -Filter { Name -eq $managerName }

# Determine Office Location based on department
$officeName = if ($OfficeLocations.ContainsKey($selectedDepartment)) { 
    $OfficeLocations[$selectedDepartment] 
} else { 
    $null 
}

# If user does not exist, create it
if ($null -eq $user) {
    try {
        $user = New-ADUser -Name "$firstName $lastName" `
                           -DisplayName "$firstName $lastName" `
                           -GivenName $firstName `
                           -Surname $lastName `
                           -SamAccountName $samAccountName `
                           -UserPrincipalName $userPrincipalName `
                           -Title $Title `
                           -Description $Title `
                           -Department $Department `
                           -Office $officeName `
                           -Company $Company `
                           -Manager $manager.DistinguishedName `
                           -EmailAddress "$userEmail" `
                           -Path $ouPath `
                           -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                           -Enabled $true `
                           -PassThru
        $outputMessages += "User $firstName $lastName created successfully in AD."
    } catch {
        $outputMessages += "Error creating user $firstName $lastName $_"
    }
}

# Add user to groups based on Title selection
$selectedGroups = $titleData | Select-Object -ExpandProperty $userInput.Title
$groupAdded = $false  # Flag to track if any group was added

foreach ($group in $selectedGroups) {
    try {
        if ($group -and (Get-ADGroup -Filter { Name -eq $group })) {
            Add-ADGroupMember -Identity $group -Members $user.DistinguishedName
            $groupAdded = $true
        }
    } catch {
    }
}

# Summarize the result
if ($groupAdded) {
    $outputMessages += "$firstName $lastName has been successfully added to the groups."
} else {
    $outputMessages += "No groups were added for $firstName $lastName."
}

# Output all collected messages at the end
$outputMessages | ForEach-Object { Write-Host $_ }