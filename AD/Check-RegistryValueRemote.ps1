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

# Get the registry path and value name from the user
$registryPath = Read-Host -Prompt "Enter the registry path"
$valueName = Read-Host -Prompt "Enter the registry value name"

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

# Run checks on remote computers asynchronously
$jobs = @()
foreach ($computer in $computers) {
    $job = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock -ArgumentList $registryPath, $valueName -AsJob
    $jobs += $job
}

# Wait for all jobs to finish and output the results
$jobs | ForEach-Object {
    $result = Receive-Job -Job $_ -Wait
    Write-Output $result
}

# Clean up the jobs
$jobs | ForEach-Object { Remove-Job -Job $_ }