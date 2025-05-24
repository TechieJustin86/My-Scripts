<#
This script is useful for administrators who need to identify GPOs with disabled settings (either user or computer) across their domain. 
It classifies GPOs based on the disabled settings and provides a clear output of which GPOs fall into each category. 
This can help in auditing GPO configurations and troubleshooting or refining group policies in a domain.
#>

# Import the GroupPolicy module
Import-Module GroupPolicy

# Get all GPOs in the domain
$gpos = Get-GPO -All

# Initialize arrays for storing GPOs with specific disabled settings
$computerDisabled = @()
$userDisabled = @()
$bothDisabled = @()

# Loop through each GPO and check its settings
foreach ($gpo in $gpos) {
    if ($gpo.GpoStatus -eq "UserSettingsDisabled") {
        $userDisabled += $gpo
    } elseif ($gpo.GpoStatus -eq "ComputerSettingsDisabled") {
        $computerDisabled += $gpo
    } elseif ($gpo.GpoStatus -eq "AllSettingsDisabled") {
        $bothDisabled += $gpo
    }
}

# Output the results
Write-Host "GPOs with Computer Settings Disabled:" -ForegroundColor Green
$computerDisabled | ForEach-Object {
    Write-Host $_.DisplayName -ForegroundColor Yellow
}

Write-Host "`nGPOs with User Settings Disabled:" -ForegroundColor Green
$userDisabled | ForEach-Object {
    Write-Host $_.DisplayName -ForegroundColor Yellow
}

Write-Host "`nGPOs with Both Settings Disabled:" -ForegroundColor Green
$bothDisabled | ForEach-Object {
    Write-Host $_.DisplayName -ForegroundColor Yellow
}

Write-Host "`nQuery Completed!" -ForegroundColor Cyan
