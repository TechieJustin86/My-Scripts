<#
This script is useful for Active Directory administrators who want to check where and when a specific user last logged on across all 
domain controllers. It ensures that the query is made to every domain controller, as the LastLogonDate attribute is not 
replicated across domain controllers and is unique to each one. This is important for accurately tracking the last logon 
activity of a user in a multi-domain controller environment.
#>

# Import the GroupPolicy module
Import-Module GroupPolicy

# Get all GPOs in the domain
$gpos = Get-GPO -All

# Initialize an array for unlinked GPOs
$unlinkedGPOs = @()

# Loop through each GPO and check its links
foreach ($gpo in $gpos) {
    # Get the links for the current GPO
    $links = Get-GPOReport -Guid $gpo.Id -ReportType XML | Out-String

    # Check if the GPO has any linked objects
    if ($links -notmatch "<Links>.*</Links>") {
        $unlinkedGPOs += $gpo
    }
}

# Output the results
if ($unlinkedGPOs.Count -gt 0) {
    Write-Host "The following GPOs are not linked to any objects:" -ForegroundColor Green
    $unlinkedGPOs | ForEach-Object {
        Write-Host $_.DisplayName -ForegroundColor Yellow
    }
} else {
    Write-Host "No unlinked GPOs found in the domain." -ForegroundColor Green
}
