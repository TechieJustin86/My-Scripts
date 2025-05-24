<#
Exports information from all users under a given OU into 3 different sheets and saves the data to C:\Temp\UserDetailsReport1.xlsx

Sheet 1 - user Details

Full Name,Email Address,Sam Account Name,Mobile Phone,IP Phone,Title,Department,Manager

Sheet 2 - User Groups
User groups that each staff membet has assigned to them

Sheet 3 - Job Title Groups
Loops thru everyones job title and pulls the groups from each.

#>

# Import necessary modules
Import-Module ActiveDirectory
Import-Module ImportExcel

# Define the OU you want to search through
$ouDN = "OU=XXX,DC=XXX,DC=XXX"

# Specify the output file
$outputFile = "C:\Temp\UserDetailsReport.xlsx"

# Ensure the output directory exists
$directory = Split-Path -Path $outputFile
if (!(Test-Path -Path $directory)) {
    New-Item -Path $directory -ItemType Directory | Out-Null
}

# Retrieve user properties from the specified OU
$users = Get-ADUser -Filter * -SearchBase $ouDN -Properties GivenName, Surname, EmailAddress, SamAccountName, MobilePhone, IPPhone, Title, Department, Manager, MemberOf |
    ForEach-Object {
        $manager = if ($_.Manager) {
            $managerDetails = Get-ADUser -Identity $_.Manager -Properties GivenName, Surname -ErrorAction SilentlyContinue
            "$($managerDetails.GivenName) $($managerDetails.Surname)"
        } else {
            "No Manager"
        }

        # Combine GivenName and Surname
        $fullName = "$($_.GivenName) $($_.Surname)"

        [PSCustomObject]@{
            FullName        = $fullName
            #GivenName       = $_.GivenName
            #Surname         = $_.Surname
            EmailAddress    = $_.EmailAddress
            SamAccountName  = $_.SamAccountName
            MobilePhone     = $_.MobilePhone
            IPPhone         = $_.IPPhone
            Title           = $_.Title
            Department      = $_.Department
            Manager         = $manager
        }
    }

# Sort the users by FullName (First Name A-Z)
$sortedUsers = $users | Sort-Object -Property FullName

# Export user details to the first sheet (UserDetails) with the combined Full Name and sorted by First Name
$sortedUsers | Export-Excel -Path $outputFile -WorksheetName "UserDetails" -AutoSize

# Prepare group membership data for the second sheet (username as header, groups under each username)
$groupMembershipData = @{}
$jobTitles = @{}

foreach ($user in $sortedUsers) {
    # Get groups for the current user
    $groups = $user.SamAccountName | ForEach-Object {
        $userGroups = Get-ADUser -Identity $_ -Properties MemberOf | Select-Object -ExpandProperty MemberOf -ErrorAction SilentlyContinue
        $userGroups | ForEach-Object { 
            ($_ -replace '^CN=([^,]+).+', '$1')
        }
    }

    # Store the groups with the username as the key
    $groupMembershipData[$user.SamAccountName] = $groups

    # Store the groups under the job title
    $jobTitle = $user.Title
    if ([string]::IsNullOrEmpty($jobTitle)) {
        $jobTitle = "No Job Title"
    }

    if (-not $jobTitles.ContainsKey($jobTitle)) {
        $jobTitles[$jobTitle] = @()
    }

    foreach ($group in $groups) {
        if ($jobTitles[$jobTitle] -notcontains $group) {
            $jobTitles[$jobTitle] += $group
        }
    }
}

# Export UserGroups sheet (usernames as headers, with groups under each)
$exportGroupData = @()
$maxGroupCount = ($groupMembershipData.Values | ForEach-Object { $_.Count }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

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
$exportJobTitleData = @()
$maxJobTitleGroupCount = ($jobTitles.Values | ForEach-Object { $_.Count }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

for ($i = 0; $i -lt $maxJobTitleGroupCount; $i++) {
    $row = @{ }

    $jobTitles.Keys | Sort-Object | ForEach-Object {
        $jobTitle = $_
        if ($jobTitles[$jobTitle].Count -gt $i) {
            $row[$jobTitle] = $jobTitles[$jobTitle][$i]
        } else {
            $row[$jobTitle] = ""
        }
    }

    $exportJobTitleData += New-Object PSObject -Property $row
}

# Export job titles and groups to the "JobTitleGroups" sheet
$exportJobTitleData | Export-Excel -Path $outputFile -WorksheetName "JobTitleGroups" -AutoSize

Write-Host "Export completed! File saved to $outputFile"