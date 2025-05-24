# Get the string we want to search for
$string = Read-Host -Prompt "What string do you want to search for?"

# Escape special regex characters in the search string
$string = [regex]::Escape($string)

# Set the domain to search for GPOs
$DomainName = $env:USERDOMAIN

# Find all GPOs in the current domain
Write-Host "Finding all the GPOs in $DomainName"
Import-Module grouppolicy

# Try to get all GPOs from the domain
try {
    $allGposInDomain = Get-GPO -All -Domain $DomainName
} catch {
    Write-Host "Error retrieving GPOs from domain: $_" -ForegroundColor Red
    exit
}

# Initialize the list of matched GPOs
$MatchedGPOList = @()

# Look through each GPO's XML for the string
Write-Host "Starting search...."
foreach ($gpo in $allGposInDomain) {
    try {
        # Get GPO report in XML format
        $report = Get-GPOReport -Guid $gpo.Id -ReportType Xml
    } catch {
        Write-Host "Error retrieving report for GPO: $($gpo.DisplayName)" -ForegroundColor Red
        continue
    }

    # Check if the report contains the search string
    if ($report -match $string) {
        Write-Host "********** Match found in: $($gpo.DisplayName) **********" -ForegroundColor Green
        $MatchedGPOList += $gpo.DisplayName
    }
}

Write-Host "`r`n"
Write-Host "Results: **************" -ForegroundColor Yellow

# Display all matched GPOs
foreach ($match in $MatchedGPOList) {
    Write-Host "Match found in: $($match)" -ForegroundColor Green
}

# If no matches were found
if ($MatchedGPOList.Count -eq 0) {
    Write-Host "No matches found in any GPOs." -ForegroundColor Red
}
