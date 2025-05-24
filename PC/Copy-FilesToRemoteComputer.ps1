<#
This script automates the process of copying two files from a server to a remote machine. It validates the existence of the files, 
creates the destination folder if necessary, and verifies the success of the copy operation. 
It's useful for file distribution tasks across multiple machines.
#>
# Define the source file paths on the server
$sourceFile1 = "PathToSourceFile1"
$sourceFile2 = "PathToSourceFile2"

# Define the remote computer name and the destination folder
$remoteComputer = Read-Host -Prompt "Enter the computer name"
$destinationFolder = "\\$remoteComputer\PathToDestinationFolder"

# Validate source files exist
if (-not (Test-Path -Path $sourceFile1)) {
    Write-Host "Source file 1 does not exist. Please verify the path." -ForegroundColor Red
    return
}

if (-not (Test-Path -Path $sourceFile2)) {
    Write-Host "Source file 2 does not exist. Please verify the path." -ForegroundColor Red
    return
}

# Check if the destination folder exists on the remote computer
if (-not (Test-Path -Path $destinationFolder)) {
    # If the folder does not exist, create it
    try {
        New-Item -Path $destinationFolder -ItemType Directory -Force
        Write-Host "The destination folder was created on the remote computer." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create the destination folder. Error: $_" -ForegroundColor Red
        return
    }
}

# Copy the first file from the server to the remote computer
try {
    Copy-Item -Path $sourceFile1 -Destination $destinationFolder -Force -ErrorAction Stop
    Write-Host "The first file was copied successfully to the remote computer." -ForegroundColor Green
}
catch {
    Write-Host "The copy operation for the first file failed. Error: $_" -ForegroundColor Red
}

# Copy the second file from the server to the remote computer
try {
    Copy-Item -Path $sourceFile2 -Destination $destinationFolder -Force -ErrorAction Stop
    Write-Host "The second file was copied successfully to the remote computer." -ForegroundColor Green
}
catch {
    Write-Host "The copy operation for the second file failed. Error: $_" -ForegroundColor Red
}

# Verify if the first file was copied successfully
$destinationFile1 = Join-Path -Path $destinationFolder -ChildPath (Split-Path -Path $sourceFile1 -Leaf)
if (Test-Path -Path $destinationFile1) {
    Write-Host "The first file was copied successfully to the remote computer." -ForegroundColor Green
} else {
    Write-Host "The copy operation for the first file failed." -ForegroundColor Red
}

# Verify if the second file was copied successfully
$destinationFile2 = Join-Path -Path $destinationFolder -ChildPath (Split-Path -Path $sourceFile2 -Leaf)
if (Test-Path -Path $destinationFile2) {
    Write-Host "The second file was copied successfully to the remote computer." -ForegroundColor Green
} else {
    Write-Host "The copy operation for the second file failed." -ForegroundColor Red
}
