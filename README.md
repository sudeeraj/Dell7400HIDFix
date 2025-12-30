# Dell Latitude 7400 HID/USB Repair Script

## Overview
This repository provides a **resume-safe PowerShell script** to fix HID/USB input issues on Dell Latitude 7400 laptops.

It performs:
- Cleanup of stale/orphaned HID devices
- Diagnostics for keyboards, mice, touchpad, and Bluetooth
- Resume-safe driver installation:
  - Chipset
  - Serial IO
  - HID Event Filter
  - Bluetooth
- Automatic reboots
- Logging and README generation

## Requirements
- Windows 10/11
- Administrator privileges
- Drivers placed in `Drivers/` folder

## Usage
1. Place the required Dell drivers in `Drivers/`:
   - Chipset
   - Serial IO
   - HID Event Filter
   - Bluetooth
2. Run PowerShell as Administrator:
   ```powershell
   cd Scripts
   .\Dell7400_HID_Fix.ps1
