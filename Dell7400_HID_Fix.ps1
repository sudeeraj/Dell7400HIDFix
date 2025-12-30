<#
.SYNOPSIS
Dell Latitude 7400 USB/HID Repair Script with Resume & Logging

.DESCRIPTION
This script performs:
1. Cleanup of stale/orphaned HID devices
2. Diagnostics of HID/USB devices (keyboard, mouse, touchpad, Bluetooth)
3. Checks for required drivers (Chipset, Serial IO, HID Event Filter, Bluetooth)
4. Provides instructions if driver files are missing
5. Installs drivers in the correct order, with automatic reboots
6. Resume-safe installation (marker file tracks progress)
7. Full logging of actions, errors, and HID verification
8. Automatic README generation for usage instructions
#>

# ======================== CONFIGURATION ========================
$DriverFolder = "C:\Dell\Drivers"
$MarkerFile   = Join-Path $DriverFolder "DriverInstallProgress.txt"
$LogFile      = Join-Path $DriverFolder "DriverInstallLog.txt"
$ReadMeFile   = Join-Path $DriverFolder "README_Dell7400_Fix.txt"

# Drivers in correct installation order
$DriverTypes = @("Chipset","Serial-IO","HID-Event-Filter","Bluetooth")

# ======================== FUNCTIONS ===========================
function LogWrite {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $LogFile -Append
    Write-Host $message
}

function Check-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Script must be run as Administrator. Exiting." -ForegroundColor Red
        exit
    }
}

function Verify-HID {
    LogWrite "=== Verifying HID/USB devices ==="
    $HIDProblems = Get-PnpDevice -Class HIDClass |
        Where-Object { $_.Status -ne "OK" -and $_.FriendlyName -notmatch "PCI Minidriver|Intel|Converted|Portable|Microsoft Input" }

    if ($HIDProblems) {
        LogWrite "Warning: Some HID devices are not OK"
        $HIDProblems | ForEach-Object {
            LogWrite "HID Device Issue - FriendlyName: $($_.FriendlyName), Status: $($_.Status), InstanceId: $($_.InstanceId)"
        }
        Write-Host "Some HID devices are not OK. Check log for details." -ForegroundColor Yellow
        return $false
    } else {
        LogWrite "All HID devices are OK ✅"
        Write-Host "All HID devices are OK ✅" -ForegroundColor Green
        return $true
    }
}

function Generate-ReadMe {
    if (-not (Test-Path $ReadMeFile)) {
        @"
Dell Latitude 7400 HID/USB Repair Script

1. Place required Dell drivers in C:\Dell\Drivers:
   - Chipset
   - Serial IO
   - HID Event Filter
   - Bluetooth

2. Run this script as Administrator.

3. Script Steps:
   a) Cleans stale/orphaned HID devices
   b) Performs full diagnostics of HID/USB devices
   c) Checks for required drivers
   d) Installs drivers in correct order
   e) Automatically reboots after each driver
   f) Resumes from last incomplete step after reboot
   g) Verifies HID/USB devices and logs results

4. Logs:
   - All actions, errors, and HID verification results are logged to DriverInstallLog.txt

5. Marker file:
   - DriverInstallProgress.txt tracks the last installed driver for resume

6. Repeat:
   - If drivers are missing, the script will instruct what to download
   - Place drivers in C:\Dell\Drivers and rerun the script

7. Tips:
   - Always run as Administrator
   - Do not move or rename driver files
   - For intermittent HID issues, re-run script; it will resume automatically

"@ | Out-File -FilePath $ReadMeFile
        LogWrite "README generated at $ReadMeFile"
    }
}

# ======================== SCRIPT START ========================
Check-Admin
Generate-ReadMe
LogWrite "===== Starting Dell Latitude 7400 HID/USB Repair Script ====="

# ======================== Step 0: Remove Stale/Orphaned HID Devices ========================
LogWrite "=== Checking for stale/orphaned HID devices ==="

$StaleHIDs = Get-PnpDevice -Class HIDClass |
             Where-Object { $_.Status -ne "OK" -and $_.FriendlyName -notmatch "PCI Minidriver|Intel|Converted|Portable|Microsoft Input" }

if ($StaleHIDs) {
    $StaleHIDs | ForEach-Object {
        LogWrite "Removing orphaned HID: $($_.FriendlyName) - InstanceId: $($_.InstanceId)"
        Disable-PnpDevice -InstanceId $_.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
        Start-Process -FilePath "pnputil.exe" -ArgumentList "/remove-device `"$($_.InstanceId)`"" -Wait
    }
    LogWrite "Stale HID cleanup completed."
} else {
    LogWrite "No stale HID devices detected."
}

# ======================== Read last completed driver index ========================
if (Test-Path $MarkerFile) {
    $LastCompletedIndex = [int](Get-Content $MarkerFile)
    LogWrite "Resuming from driver index $LastCompletedIndex"
} else {
    $LastCompletedIndex = -1
    LogWrite "Starting fresh installation sequence"
}

# ======================== Step 1: Diagnostics ========================
$hid_ok = Verify-HID
if ($hid_ok -and $LastCompletedIndex -eq 3) {
    LogWrite "All drivers installed and HID devices OK. No action required."
    exit
}

# ======================== Step 2: Driver installation loop ========================
for ($i = $LastCompletedIndex + 1; $i -lt $DriverTypes.Count; $i++) {
    $type = $DriverTypes[$i]
    $driver = Get-ChildItem -Path $DriverFolder -Filter "*.EXE" |
              Where-Object { $_.Name -match $type } |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

    if ($driver) {
        LogWrite "Installing driver: $($driver.Name)"
        try {
            Start-Process -FilePath $driver.FullName -ArgumentList "/s /quiet /norestart" -Wait
            LogWrite "Successfully installed $($driver.Name)"
        } catch {
            LogWrite "Error installing $($driver.Name): $_"
            Start-Process -FilePath $driver.FullName -Wait
        }

        # Update marker file
        Set-Content -Path $MarkerFile -Value $i

        # Reboot after each driver
        LogWrite "Rebooting to apply changes..."
        Restart-Computer -Force
        break
    } else {
        LogWrite "Driver for type '$type' not found. Please download the latest $type driver from Dell Support and place it in C:\Dell\Drivers."
        Write-Host "Driver for type '$type' missing. Place the correct driver in C:\Dell\Drivers and rerun script." -ForegroundColor Yellow
        exit
    }
}

# ======================== Step 3: After installation ========================
if ($i -ge $DriverTypes.Count) {
    if (Test-Path $MarkerFile) { Remove-Item $MarkerFile }
    LogWrite "All drivers installed successfully!"

    # Step 4: Verify HID devices again
    Verify-HID
}
