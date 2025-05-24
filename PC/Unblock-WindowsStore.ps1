# Function to unblock the Windows Store by modifying registry settings
function Unblock-WindowsStore {
    # Define the registry path and key names
    $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"
    $RemovestoreRegKey = "RemoveWindowsStore"
    $DisableStore = "DisableStoreApps"

    # Check if the registry path exists
    if (Test-Path $regPath) {
        # Set the registry keys to unblock the Windows Store
        Set-ItemProperty -Path $regPath -Name $RemovestoreRegKey -Value 0
        Set-ItemProperty -Path $regPath -Name $DisableStore -Value 0
        Write-Host "Windows Store has been unblocked." -ForegroundColor Green
    } else {
        Write-Host "Registry entry for blocking Windows Store does not exist." -ForegroundColor Yellow
    }
}

# Call the function to unblock the Windows Store
Unblock-WindowsStore