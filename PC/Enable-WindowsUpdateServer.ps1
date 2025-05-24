# Set execution policy to RemoteSigned (if needed)
Set-ExecutionPolicy RemoteSigned -Force

# Define the registry path and key
$regPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
$regKeyName = "UseWUServer"
$regValue = 1

# Check if the registry path exists
if (Test-Path $regPath) {
    # Set the registry value to disable using a Windows Update server
    Set-ItemProperty -Path $regPath -Name $regKeyName -Value $regValue
    Write-Host "Windows Update server usage has been disabled." -ForegroundColor Green
} else {
    Write-Host "Registry path for Windows Update AU does not exist." -ForegroundColor Yellow
}
