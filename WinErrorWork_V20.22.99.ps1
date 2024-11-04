<#
.SYNOPSIS
This script fixes Windows Update errors, specifically error 0x80246007, by performing several maintenance tasks, including disabling and restarting key services, renaming system folders, and removing leftover 'pending.xml' files.

.DESCRIPTION
The script performs the following steps to resolve Windows Update issues:

1. **Logging**: The script generates a log file to record all actions performed during the script execution.
2. **Service Management**: It disables services (BITS, wuauserv, cryptsvc), records their original startup types, stops them, and later restores them to their original states.
3. **Folder Renaming**: The script renames critical folders like `SoftwareDistribution` and `Catroot2` by adding versioning (`_01`, `_02`, etc.) to preserve historical data. Older versions are deleted, and the base folder is renamed.
4. **File Cleanup**: The script removes files ending with `pending.xml` from the `WinSxS` folder to clear pending update statuses.
5. **Service Restoration**: After performing the necessary updates, the script re-enables the services and restores their original startup settings.
6. **Logging**: Logs each action taken by the script to a timestamped log file for troubleshooting and audit purposes.

.PARAMETER folderPath
The path to the folder to be renamed, used in the folder renaming function.

.PARAMETER maxSuffix
The maximum suffix value (e.g., _01, _02, _03) for versioning folder names.

.EXAMPLE
.\Fix-WindowsUpdate.ps1
This example runs the script to fix Windows Update error 0x80246007.

.NOTES
File Name      : Fix-WindowsUpdate.ps1
Author         : HarshBharath
Date           : 13-SEPT-2024
PowerShell Version : 5.1 or higher
Requires       : Administrator Privileges
#>

# Run as Administrator 

# Generate a unique log file name with timestamp
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFile = "C:\WindowsUpdateFixLog_$timestamp.txt"

# Function to log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Type] - $Message"
    Write-Host $logMessage
    Add-Content -Path $LogFile -Value $logMessage
}

# Function to change service startup type using sc config
function Set-ServiceStartupType {
    param (
        [string]$ServiceName,
        [string]$StartupType
    )

    # Map the startup type to sc config acceptable values
    $startupTypeMapping = @{
        "Auto"           = "auto"
        "DelayedAuto"    = "delayed-auto"
        #"Demand"         = "demand"
        "Manual"         = "demand"
        "Disabled"       = "disabled"
    }

    try {
        if ($startupTypeMapping.ContainsKey($StartupType)) {
            $mappedType = $startupTypeMapping[$StartupType]
        } else {
            Write-Log "Invalid startup type '$StartupType' for service $ServiceName. Valid types are: Auto, DelayedAuto, Demand, Disabled." "ERROR"
            return
        }

        # Run the sc config command to change the service startup type
        $command = "config $ServiceName start=$mappedType"
        Start-Process -FilePath "sc.exe" -ArgumentList $command -NoNewWindow -Wait
        Write-Log "Changed startup type of $ServiceName to $mappedType" "INFO"
    } catch {
        Write-Log "Failed to change startup type of $ServiceName to $mappedType. Error: $_" "ERROR"
    }
}

# Function to rename folder with backup strategy
function Rename-FolderWithReverseSuffixForce {
    param (
        [string]$folderPath,
        [int]$maxSuffix = 5  # Maximum suffix limit (e.g., _01 to _05)
    )

    # Check if the folder exists
    if (Test-Path -Path $folderPath -PathType Container) {
        # Construct the folder path with the maximum suffix (_05)
        $maxFolderPath = "$folderPath`_{0:D2}" -f $maxSuffix

        # If the folder with the max suffix exists, delete it
        if (Test-Path -Path $maxFolderPath -PathType Container) {
            Remove-Item -Path $maxFolderPath -Recurse -Force
            Write-Log "Folder '$maxFolderPath' has been deleted."
        }

        # Start from the maxSuffix and work backward
        for ($i = $maxSuffix - 1; $i -ge 1; $i--) {
            $currentSuffix = "{0:D2}" -f $i  # Format the number as _01, _02, etc.
            $currentFolderPath = "$folderPath`_$currentSuffix"

            # If a folder with the current suffix exists, rename it to the next suffix
            if (Test-Path -Path $currentFolderPath -PathType Container) {
                $nextSuffix = "{0:D2}" -f ($i + 1)
                $nextFolderPath = "$folderPath`_$nextSuffix"
                Rename-Item -Path $currentFolderPath -NewName $nextFolderPath
                Write-Log "Folder '$currentFolderPath' renamed to '$nextFolderPath'."
            }
        }

        # Finally, rename the original folder to _01
        $newFolderPath = "$folderPath`_01"
        Rename-Item -Path $folderPath -NewName $newFolderPath
        Write-Log "Folder '$folderPath' has been renamed to '$newFolderPath'."
    } else {
        Write-Log "Folder '$folderPath' does not exist."
    }
}

# Start logging
Write-Log "Starting script to fix Windows Update error 0x80246007"

# Step 1: Check and record the startup type of services
Write-Log "Step 1: Checking and recording the startup type of services..."
$services = @("BITS", "wuauserv", "cryptsvc")
$serviceStartupTypes = @{}

foreach ($service in $services) {
    $serviceInfo = Get-WmiObject -Class Win32_Service -Filter "Name='$service'"
    $serviceStartupTypes[$service] = $serviceInfo.StartMode
    Write-Log "Recorded startup type of $service : $($serviceStartupTypes[$service])"
}

# Step 2: Disable the services
Write-Log "Step 2: Disabling services..."
foreach ($service in $services) {
    Set-ServiceStartupType -ServiceName $service -StartupType "Disabled"
}

# Step 3: Stop the services
Write-Log "Step 3: Stopping services..."
foreach ($service in $services) {
    try {
        Stop-Service -Name $service -Force -ErrorAction Stop
        Write-Log "Successfully stopped $service."
    } catch {
        Write-Log "Failed to stop $service. Error: $_" "ERROR"
    }
}

# Step 4: Complete the activities (Deleting or renaming files, etc.)
Write-Log "Step 4: Performing update activities..."

# Rename SoftwareDistribution folder
Write-Log "Renaming SoftwareDistribution folder..."
$SoftwareDistributionFolder = "$env:windir\SoftwareDistribution"
Rename-FolderWithReverseSuffixForce -folderPath $SoftwareDistributionFolder -maxSuffix 5

# Rename Catroot2 folder
Write-Log "Renaming Catroot2 folder..."
$Catroot2Folder = "$env:windir\System32\catroot2"
Rename-FolderWithReverseSuffixForce -folderPath $Catroot2Folder -maxSuffix 5

# Remove files ending with 'pending.xml'
Write-Log "Removing files that end with 'pending.xml'..."
try {
    $PendingXmlFiles = Get-ChildItem -Path "$env:windir\WinSxS\" -Filter "*pending.xml" -ErrorAction Stop
    foreach ($file in $PendingXmlFiles) {
        Remove-Item -Path $file.FullName -Force -ErrorAction Stop
        Write-Log "Removed $($file.FullName)"
    }
} catch {
    Write-Log "Failed to remove pending.xml files. Error: $_" "ERROR"
}

# Step 5: Restore the original startup types
Write-Log "Step 5: Restoring the original startup types of services..."
foreach ($service in $services) {
    # Wait for a short period to ensure previous operations are complete
    Start-Sleep -Seconds 5
    Set-ServiceStartupType -ServiceName $service -StartupType $serviceStartupTypes[$service]
}

# Step 6: Start the services
Write-Log "Step 6: Starting services..."
foreach ($service in $services) {
    try {
        Start-Service -Name $service -ErrorAction Stop
        Write-Log "Successfully started $service."
    } catch {
        Write-Log "Failed to start $service. Error: $_" "ERROR"
    }
}

Write-Log "Script execution completed. Please restart your system if necessary."
