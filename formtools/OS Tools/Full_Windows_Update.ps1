# Add assembly for MessageBox
Add-Type -AssemblyName System.Windows.Forms

# Function to check and install required module
Function Check-Module {
    $module = 'PSWindowsUpdate'
    if (-Not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -SkipPublisherCheck
    }
    Import-Module $module
}

# Run as administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Run this script as an Administrator." -ForegroundColor Red
    Exit
}

# Check and import required module
Check-Module

# Initialize reboot flag
$rebootRequired = $false

# Function to install updates and check for reboot requirement
Function Install-Updates {
    Write-Host "Starting Windows Update..." -ForegroundColor Green
    $updates = Get-WUInstall -MicrosoftUpdate -IgnoreUserInput -AcceptAll -IgnoreReboot
    if ($updates | Where-Object { $_.RebootRequired }) {
        $global:rebootRequired = $true
    }
}

# Function to install optional updates
Function Install-OptionalUpdates {
    Write-Host "Installing Optional Updates..." -ForegroundColor Green
    $updates = Get-WUInstall -NotCategory "Critical Updates","Security Updates" -AcceptAll -IgnoreReboot
    if ($updates | Where-Object { $_.RebootRequired }) {
        $global:rebootRequired = $true
    }
}

# Function to install feature updates
Function Install-FeatureUpdates {
    Write-Host "Installing Feature Updates..." -ForegroundColor Green
    # Replace this part with the appropriate cmdlet or method to install feature updates
}

# Run the functions
Install-Updates
Install-OptionalUpdates
Install-FeatureUpdates

# Notify about reboot if required
if ($rebootRequired) {
    $result = [System.Windows.Forms.MessageBox]::Show("Updates installed. A reboot is required. Would you like to reboot now?", "Reboot Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -eq "Yes") {
        Restart-Computer
    }
} else {
    Write-Host "Windows Update process completed." -ForegroundColor Green
}
