# Import the Active Directory module
Import-Module ActiveDirectory

# Function to get all OUs and let the user select one
function Select-OU {
    $ous = Get-ADOrganizationalUnit -Filter * | Select-Object -Property DistinguishedName, Name
    $ouList = $ous | ForEach-Object { "$($_.DistinguishedName) - $($_.Name)" }

    $selectedOU = $null
    while (-not $selectedOU) {
        $selectedOU = $ouList | Out-GridView -Title "Select an OU" -OutputMode Single
    }

    return $ous | Where-Object { $_.DistinguishedName -eq $selectedOU.Split(' - ')[0] }
}

# Define the registry path and the value name
$registryPath = "HKLM:\SOFTWARE\"
$valueName = "XXX"

# Get the selected OU
$selectedOU = Select-OU
Write-Output "Selected OU: $($selectedOU.DistinguishedName)"

# Get a list of all computers in the selected OU
$computers = Get-ADComputer -Filter * -SearchBase $selectedOU.DistinguishedName | Select-Object -ExpandProperty Name

# Define a script block to run on each computer
$scriptBlock = {
    param ($registryPath, $valueName)

    try {
        $registryValue = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction Stop
        return "Computer: $env:COMPUTERNAME - Registry value '$valueName': $($registryValue.$valueName)"
    } catch {
        return "Computer: $env:COMPUTERNAME - Registry value '$valueName' not found in '$registryPath'."
    }
}

# Iterate over each computer and check the registry value
foreach ($computer in $computers) {
    try {
        # Use Invoke-Command to run the script block on each remote computer
        $result = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock -ArgumentList $registryPath, $valueName -ErrorAction Stop
        Write-Output $result
    } catch {
        Write-Output "Computer: $computer - Unable to connect or access registry."
    }
}