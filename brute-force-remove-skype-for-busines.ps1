# Ensure the script is run as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as an administrator."
    exit
}

# Function to remove registry keys
function Remove-RegistryKey {
    param (
        [string]$path
    )
    try {
        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
        Write-Host "Removed registry key: $path"
    } catch {
        Write-Host "Failed to remove registry key: $path. Error: $_"
    }
}

# Uninstall Skype for Business
$skypeProcess = Get-Process -Name "lync" -ErrorAction SilentlyContinue
if ($skypeProcess) {
    Stop-Process -Name "lync" -Force
}

# Define paths to delete
$skypeFolders = @(
    "$env:ProgramFiles\Microsoft Office\Office16\Skype for Business"
    "$env:APPDATA\Microsoft\Office\16.0\Skype for Business"
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Skype for Business"
)

# Delete Skype for Business folders
foreach ($folder in $skypeFolders) {
    if (Test-Path $folder) {
        Remove-Item -Path $folder -Recurse -Force
        Write-Host "Deleted folder: $folder"
    } else {
        Write-Host "Folder not found: $folder"
    }
}

# Remove registry keys
$registryKeys = @(
    "HKCU:\Software\Microsoft\Communicator",
    "HKCU:\Software\Microsoft\Office\16.0\Lync",
    "HKLM:\Software\Microsoft\Office\16.0\Lync",
    "HKLM:\Software\Wow6432Node\Microsoft\Office\16.0\Lync"
)

foreach ($key in $registryKeys) {
    Remove-RegistryKey -path $key
}

# Remove the Skype for Business installer entry from the Programs and Features
try {
    $skypeUninstallKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    $skypeUninstallKeyX64 = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    
    $skypeUninstallKeys = Get-ChildItem -Path $skypeUninstallKey | Where-Object { $_.GetValue("DisplayName") -like "*Skype for Business*" }
    $skypeUninstallKeysX64 = Get-ChildItem -Path $skypeUninstallKeyX64 | Where-Object { $_.GetValue("DisplayName") -like "*Skype for Business*" }

    foreach ($key in $skypeUninstallKeys) {
        Remove-RegistryKey -path $key.PSPath
    }

    foreach ($key in $skypeUninstallKeysX64) {
        Remove-RegistryKey -path $key.PSPath
    }
} catch {
    Write-Host "Failed to remove Skype for Business uninstall entry. Error: $_"
}

Write-Host "Skype for Business has been forcefully removed."