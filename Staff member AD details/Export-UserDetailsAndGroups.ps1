Import-Module ActiveDirectory
Import-Module ImportExcel

# Define the OU you want to search through
$ouDN = "OU=XXX,DC=XXX,DC=org"

# Specify the output file
$outputFile = "C:\Temp\UserDetlsReport.xlsx"

# Ensure the output directory exists
$directory = Split-Path -Path $outputFile
if (!(Test-Path -Path $directory)) {
    New-Item -Path $directory -ItemType Directory | Out-Null
}

#region UserDetails

# Retrieve user properties from the specified OU
$users = Get-ADUser -Filter * -SearchBase $ouDN -Properties GivenName, Surname, EmailAddress, SamAccountName, MobilePhone, IPPhone, Title, Department, Manager, MemberOf

# Create an array of user objects
$users = $users | ForEach-Object {
    try {
        # Resolve the user's manager if available
        $manager = if ($_.Manager) {
            $managerDetails = Get-ADUser -Identity $_.Manager -Properties GivenName, Surname -ErrorAction SilentlyContinue
            "$($managerDetails.GivenName) $($managerDetails.Surname)"
        } else {
            "No Manager"
        }

        # Combine GivenName and Surname for Full Name
        $fullName = "$($_.GivenName) $($_.Surname)"
        
        # Return the user data as a custom object
        [PSCustomObject]@{
            GivenName       = $_.GivenName
            Surname         = $_.Surname
            FullName        = $fullName
            EmailAddress    = if ($_.EmailAddress) { $_.EmailAddress } else { "No Email" }
            SamAccountName  = $_.SamAccountName
            MobilePhone     = if ($_.MobilePhone) { $_.MobilePhone } else { "N/A" }
            IPPhone         = if ($_.IPPhone) { $_.IPPhone } else { "N/A" }
            Title           = if ($_.Title) { $_.Title } else { "No Title" }
            Department      = if ($_.Department) { $_.Department } else { "No Department" }
            Manager         = $manager
            SID             = $_.SID.Value
            MemberOf        = ($_.MemberOf | ForEach-Object { (Get-ADGroup -Identity $_).Name }) -join ", "
        }
    } catch {
        Write-Warning "Error processing user $_.SamAccountName: $_"
        return $null
    }
}

# Sort the users by FullName (First Name A-Z)
$sortedUsers = $users | Sort-Object -Property FullName

# Export user details to the first sheet (UserDetails) with the combined Full Name and sorted by First Name
$sortedUsers | Export-Excel -Path $outputFile -WorksheetName "UserDetails" -AutoSize -BoldTopRow

# Prepare group membership data for the second sheet (username as header, groups under each username)
$groupMembershipData = @{}

foreach ($user in $sortedUsers) {
    try {
        # Get groups for the current user
        $userGroups = Get-ADUser -Identity $user.SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf -ErrorAction SilentlyContinue
        $groups = $userGroups | ForEach-Object { ($_ -replace '^CN=([^,]+).+', '$1') }

        # Store the groups with the username as the key
        $groupMembershipData[$user.SamAccountName] = $groups
    } catch {
        Write-Warning "Error retrieving groups for user $($user.SamAccountName): $_"
    }
}

# Export UserGroups sheet (usernames as headers, with groups under each)
$exportGroupData = @()
$maxGroupCount = ($groupMembershipData.Values | ForEach-Object { $_.Count }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

#endregion

#region UserGroups
# Ensure that the group membership sheet has equal columns for each user
for ($i = 0; $i -lt $maxGroupCount; $i++) {
    $row = @{ }

    $groupMembershipData.Keys | Sort-Object | ForEach-Object {
        $username = $_
        if ($groupMembershipData[$username].Count -gt $i) {
            $row[$username] = $groupMembershipData[$username][$i]
        } else {
            $row[$username] = ""
        }
    }

    $exportGroupData += New-Object PSObject -Property $row
}

# Export group memberships to the "UserGroups" sheet
$exportGroupData | Export-Excel -Path $outputFile -WorksheetName "UserGroups" -AutoSize

# Prepare Job Title groups sheet (job titles as headers)
$jobTitleGroups = @{}
$maxJobTitleGroupCount = 0

foreach ($user in $sortedUsers) {
    try {
        $jobTitle = $user.Title
        if ([string]::IsNullOrEmpty($jobTitle)) {
            $jobTitle = "No Job Title"
        }

        # Ensure job titles are grouped
        if (-not $jobTitleGroups.ContainsKey($jobTitle)) {
            $jobTitleGroups[$jobTitle] = @()
        }

        $userGroups = $groupMembershipData[$user.SamAccountName]
        $jobTitleGroups[$jobTitle] += $userGroups
    } catch {
        Write-Warning "Error processing job title for user $($user.SamAccountName): $_"
    }
}

# Export Job Title groups sheet (group memberships under each job title)
$exportJobTitleData = @()
$maxJobTitleGroupCount = ($jobTitleGroups.Values | ForEach-Object { $_.Count }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

# Ensure equal columns for job title groups
for ($i = 0; $i -lt $maxJobTitleGroupCount; $i++) {
    $row = @{ }

    $jobTitleGroups.Keys | Sort-Object | ForEach-Object {
        $jobTitle = $_
        if ($jobTitleGroups[$jobTitle].Count -gt $i) {
            $row[$jobTitle] = $jobTitleGroups[$jobTitle][$i]
        } else {
            $row[$jobTitle] = ""
        }
    }

    $exportJobTitleData += New-Object PSObject -Property $row
}

#endregion

# Export job titles and groups to the "JobTitleGroups" sheet
$exportJobTitleData | Export-Excel -Path $outputFile -WorksheetName "JobTitleGroups" -AutoSize

Write-Host "Export completed! File saved to $outputFile"
