# Define the registry key path and the value name
$registryPath = "HKLM:\SOFTWARE\"
$valueName = "XXX"

# Retrieve the registry value
try {
    $registryValue = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction Stop

    # Output the value
    Write-Output "Registry value '$valueName' in '$registryPath': $($registryValue.$valueName)"
} catch {
    # Handle the case where the registry value does not exist
    Write-Output "Registry value '$valueName' not found in '$registryPath'."
}