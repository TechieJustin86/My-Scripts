<#
This script enables administrators to list installed Appx packages and selectively remove them based on user input. 
It provides error handling for invalid input and protected system apps that cannot be removed on a per-user basis. 
Custom error messages are displayed for clarity and troubleshooting.
#>
Import-LocalizedData -BindingVariable Messages

# Function to create a custom error record
Function PSCustomErrorRecord {
    Param (
        [Parameter(Mandatory=$true, Position=1)] [String]$ExceptionString,
        [Parameter(Mandatory=$true, Position=2)] [String]$ErrorID,
        [Parameter(Mandatory=$true, Position=3)] [System.Management.Automation.ErrorCategory]$ErrorCategory,
        [Parameter(Mandatory=$true, Position=4)] [PSObject]$TargetObject
    )
    Process {
        $exception = New-Object System.Management.Automation.RuntimeException($ExceptionString)
        $customError = New-Object System.Management.Automation.ErrorRecord($exception, $ErrorID, $ErrorCategory, $TargetObject)
        return $customError
    }
}

# Function to remove specified Appx packages
Function RemoveAppxPackage {
    $index = 1
    $apps = Get-AppxPackage

    # Display app list with IDs
    Write-Host "ID`t App name"
    foreach ($app in $apps) {
        Write-Host " $index`t $($app.name)"
        $index++
    }
    
    Do {
        $IDs = Read-Host -Prompt "Which Apps do you want to remove? `nInput their IDs, separated by commas"
    } While ($IDs -eq "")

    # Validate input
    try {
        [int[]]$IDs = $IDs -split ","
    }
    catch {
        $errorMsg = $Messages.IncorrectInput
        $errorMsg = $errorMsg -replace "Placeholder01", $IDs
        $customError = PSCustomErrorRecord -ExceptionString $errorMsg -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $pscmdlet
        $pscmdlet.WriteError($customError)
        return
    }

    # Process each ID
    foreach ($ID in $IDs) {
        if ($ID -ge 1 -and $ID -le $apps.Count) {
            $ID--
            $AppName = $apps[$ID].name

            # Attempt to remove the app
            Remove-AppxPackage -Package $apps[$ID] -ErrorAction SilentlyContinue
            if (-not(Get-AppxPackage -Name $AppName)) {
                Write-Host "$AppName has been removed successfully" -ForegroundColor Green
            } else {
                Write-Warning "Remove '$AppName' failed! This app is part of Windows and cannot be uninstalled on a per-user basis."
            }
        }
        else {
            # Handle invalid IDs
            $errorMsg = $Messages.WrongID
            $errorMsg = $errorMsg -replace "Placeholder01", $ID
            $customError = PSCustomErrorRecord -ExceptionString $errorMsg -ErrorCategory NotSpecified -ErrorID 1 -TargetObject $pscmdlet
            $pscmdlet.WriteError($customError)
        }
    }
}

# Call the function to remove the specified Appx packages
RemoveAppxPackage
