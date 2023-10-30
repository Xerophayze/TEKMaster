<#
    .SYNOPSIS
    A PowerShell script to display basic system information in a graphical user interface (GUI).

    .DESCRIPTION
    This GUI-based PowerShell script displays:
    1. Operating System (OS) information.
    2. Network configuration and status.
    3. Storage details, including S.M.A.R.T. health status for drives.
    
    The interface presents information in a grid layout and provides options to save the gathered information as an HTML file or fetch S.M.A.R.T. details for storage devices. Additionally, there's an option to open the Event Viewer directly from the interface.

    .AUTHOR
    Eric Thorup

    .COPYRIGHT
    Copyright (c) 2023 TEK Utah LLC. All rights reserved.
#>

# System information
# This script will display a form that has basic system information laid out in a grid form.

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
	public class ConsoleWindow {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@ -Language CSharp

# Get the current script or exe's directory
$currentDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Information"
$form.Size = New-Object System.Drawing.Size(1024,512)
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Add_Activated({
    $form.TopMost = $true
})


function Get-OSInfo {
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $computerInfo = Get-CimInstance Win32_ComputerSystem

    $computerName = $env:COMPUTERNAME
    $windowsVersion = $osInfo.Caption
    $osArchitecture = $osInfo.OSArchitecture
    $logicalProcessors = $computerInfo.NumberOfLogicalProcessors
    $physicalMemory = [math]::Round($computerInfo.TotalPhysicalMemory / 1GB, 2)
    $lastBoot = $osInfo.LastBootUpTime
    $osBuild = $osInfo.BuildNumber
    $osVersion = $osInfo.Version
    $systemManufacturer = $computerInfo.Manufacturer
    $systemModel = $computerInfo.Model
    $username = $env:USERNAME

    return @"
    
Computer Name: $computerName
Windows Version: $windowsVersion
OS Architecture: $osArchitecture
Logical Processors: $logicalProcessors
Total Physical Memory (GB): $physicalMemory
Last Boot-Up Time: $lastBoot
OS Build Number: $osBuild
OS Version: $osVersion
System Manufacturer: $systemManufacturer
System Model: $systemModel
Username: $username
"@
}

function SaveSmartInfoToFile {
    $drives = Get-CimInstance Win32_DiskDrive

    $allSmartDetails = @()

    foreach ($drive in $drives) {
        # Convert Windows drive name to Linux style for smartctl
        $driveIndex = $drive.DeviceID -replace '\\\\.\\PHYSICALDRIVE', ''
        $driveNameLinuxStyle = "/dev/sd" + [char]([int][char]'a' + [int]$driveIndex)

        # Fetch S.M.A.R.T. details
        $smartDetails = & "$PSScriptRoot\smartmontools\smartctl.exe" -x $driveNameLinuxStyle
        $allSmartDetails += $smartDetails
        $allSmartDetails += "=================================================="
    }

# Create a SaveFileDialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save S.M.A.R.T. Info As"
    $saveFileDialog.FileName = "$env:COMPUTERNAME SMART Info.txt"

    # Show the SaveFileDialog and check if the user clicked OK
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $allSmartDetails | Out-File $saveFileDialog.FileName
    }
}

function Get-NetworkInfo {
    $allAdapters = Get-WmiObject Win32_NetworkAdapter
    $networkConfigurations = Get-WmiObject Win32_NetworkAdapterConfiguration | Group-Object -Property Index

    $networkInfoText = @()

    foreach ($adapter in $allAdapters) {
        $adapterConfig = $networkConfigurations | Where-Object { $_.Name -eq $adapter.Index } | Select-Object -ExpandProperty Group

        if ($adapterConfig.IPAddress) {
            $ipAddress = $adapterConfig.IPAddress[0]
            $subnet = $adapterConfig.IPSubnet[0]
            $gateway = $adapterConfig.DefaultIPGateway -join ', '
            $dnsServers = $adapterConfig.DNSServerSearchOrder -join ', '
            $description = $adapter.Description

            $adapterInfo = @"
            
Adapter Description: $description
IP Address: $ipAddress
Subnet: $subnet
Gateway: $gateway
DNS Servers: $dnsServers
"@
            $networkInfoText += $adapterInfo
        } elseif ($adapter.NetConnectionStatus -eq 7) {
            $description = $adapter.Description
            $adapterInfo = @"
            
Adapter Description: $description
Status: Disabled
"@
            $networkInfoText += $adapterInfo
        }
    }

    return $networkInfoText -join "`n"
}

function Get-StorageInfo {
    $drives = Get-CimInstance Win32_DiskDrive
	
    $storageInfoText = @()

    foreach ($drive in $drives) {
        $sizeGB = [math]::Round($drive.Size / 1GB, 2)
        
        $partitions = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($drive.DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition"
        
        $volumes = $partitions | ForEach-Object { Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($_.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition" }
        
        $totalSpaceGB = [math]::Round(($volumes.Size | Measure-Object -Sum).Sum / 1GB, 2)
        $freeSpaceGB = [math]::Round(($volumes.FreeSpace | Measure-Object -Sum).Sum / 1GB, 2)
        $usedSpaceGB = $totalSpaceGB - $freeSpaceGB

        # Convert Windows drive name to Linux style for smartctl
        $driveIndex = $drive.DeviceID -replace '\\\\.\\PHYSICALDRIVE', ''
        $driveNameLinuxStyle = "/dev/sd" + [char]([int][char]'a' + [int]$driveIndex)

        # Fetch S.M.A.R.T. details
		# First, determine the path for the smartmontools directory:
		$smartmontoolsPath = Join-Path -Path $PSScriptRoot -ChildPath "smartmontools"

		# Then, determine the path for the smartctl.exe inside the smartmontools directory:
		$smartctlPath = Join-Path -Path $smartmontoolsPath -ChildPath "smartctl.exe"

		# Fetch S.M.A.R.T. details using the resolved path:
		$smartDetails = & $smartctlPath -x $driveNameLinuxStyle

        # Extract overall health status
        $healthStatusLine = $smartDetails | Where-Object { $_ -like "*SMART overall-health self-assessment test result:*" }

        if ($healthStatusLine) {
            $healthStatus = ($healthStatusLine -split ":")[1].Trim()
        } else {
            $healthStatus = "Unknown"
        }

        $driveInfo = @"
        
		
Model: $($drive.Model)
Serial Number: $($drive.SerialNumber)
Capacity: ${sizeGB}GB
Used Space: ${usedSpaceGB}GB
Free Space: ${freeSpaceGB}GB
S.M.A.R.T. Health Status: $healthStatus
"@

        $storageInfoText += $driveInfo
    }

    return $storageInfoText -join "`n"
}

# Create a TableLayoutPanel to manage the layout
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.RowCount = 3
$tableLayoutPanel.ColumnCount = 3
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$form.Controls.Add($tableLayoutPanel) | Out-Null

for ($localY = 0; $localY -lt 3; $localY++) {
    for ($localX = 0; $localX -lt 3; $localX++) {
        $textBoxNumber = $localX + 1 + ($localY * 3)

        # Create panel for label and textbox
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = [System.Windows.Forms.DockStyle]::Fill

        # Create label
        $label = New-Object System.Windows.Forms.Label
        $label.Dock = [System.Windows.Forms.DockStyle]::Top
        $label.Height = 15
        $label.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0) # Add margin at the bottom

        switch ($textBoxNumber) {
            1 { $label.Text = "OS Info" }
            2 { $label.Text = "Network Info" }
            default { $label.Text = "Info $textBoxNumber" }
        }

        $panel.Controls.Add($label) | Out-Null

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $textBox.Multiline = $true
        $textBox.ReadOnly = $true
        $textBox.Name = "TextBox$textBoxNumber"
        $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $textBox.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 0) # Add margin at the top

        if ($textBoxNumber -eq 1) {
            $textBox.Text = Get-OSInfo
        } elseif ($textBoxNumber -eq 2) {
            $textBox.Text = Get-NetworkInfo
        } elseif ($textBoxNumber -eq 3) {
            $textBox.Text = Get-StorageInfo
            $label.Text = "Storage Info"
        }

        $panel.Controls.Add($textBox) | Out-Null
        $tableLayoutPanel.Controls.Add($panel, $localX, $localY) | Out-Null
    }
}

# Create a FlowLayoutPanel for buttons
$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$buttonPanel.Height = 40
$form.Controls.Add($buttonPanel) | Out-Null

# Create the Save button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Height = 30
$saveButton.Add_Click({
    $htmlContent = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <title>System Information</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <style type="text/css">
        .theader {background-color: lime; text-align: left;}
        h3 {font-size:18pt; font-weight:bold; margin-bottom: 2px;}
        tr:hover td { background: #ddd; }
    </style>
</head>
<body bgcolor="#FFFFFF" TEXT="#000000">
    <h3 align="center">System Information</h3>
    <table width="100%">
        <tr valign=middle>
            <td class="theader">OS Info</td>
            <td>$(($form.Controls[0].Controls[0].Controls[1]).Text.Replace("`r`n", "<br/>"))</td>
        </tr>
        <tr valign=middle>
            <td class="theader">Network Info</td>
            <td>$(($form.Controls[0].Controls[1].Controls[1]).Text.Replace("`r`n", "<br/>"))</td>
        </tr>
        <tr valign=middle>
            <td class="theader">Storage Info</td>
            <td>$(($form.Controls[0].Controls[2].Controls[1]).Text.Replace("`r`n", "<br/>"))</td>
        </tr>
    </table>
</body>
</html>
"@
    # Create a SaveFileDialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save System Information As"
    $saveFileDialog.FileName = "systeminfo.html"

    # Show the SaveFileDialog and check if the user clicked OK
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $htmlContent | Out-File $saveFileDialog.FileName
    }
})

# Create the Save S.M.A.R.T. Info button
$saveSmartInfoButton = New-Object System.Windows.Forms.Button
$saveSmartInfoButton.Text = "Save S.M.A.R.T. Info"

# Measure text width
$graphics = $saveSmartInfoButton.CreateGraphics()
$size = $graphics.MeasureString($saveSmartInfoButton.Text, $saveSmartInfoButton.Font)

# Set button size
$padding = 10 # Adjust as necessary for padding on both sides
$saveSmartInfoButton.Height = 30
$saveSmartInfoButton.Width = [System.Math]::Ceiling($size.Width) + ($padding * 2)

$saveSmartInfoButton.Add_Click({ SaveSmartInfoToFile })

# Create the Open Event Viewer button
$openEventViewerButton = New-Object System.Windows.Forms.Button
$openEventViewerButton.Text = "Open Event Viewer"
$openEventViewerButton.Height = 30

# Measure text width for the Open Event Viewer button
$graphics = $openEventViewerButton.CreateGraphics()
$size = $graphics.MeasureString($openEventViewerButton.Text, $openEventViewerButton.Font)

# Set button size for the Open Event Viewer button
$padding = 10 # Adjust as necessary for padding on both sides
$openEventViewerButton.Width = [System.Math]::Ceiling($size.Width) + ($padding * 2)

$openEventViewerButton.Add_Click({
    # Temporarily remove the TopMost setting from the main form
    $form.TopMost = $false

    # Start the Event Viewer
    Start-Process "eventvwr.msc"
})

# Add all buttons to the FlowLayoutPanel
$buttonPanel.Controls.Add($saveButton) | Out-Null
$buttonPanel.Controls.Add($saveSmartInfoButton) | Out-Null
$buttonPanel.Controls.Add($openEventViewerButton) | Out-Null

# Minimize the PowerShell console window
$consoleWindowHandle = [ConsoleWindow]::GetConsoleWindow()
[ConsoleWindow]::ShowWindow($consoleWindowHandle, 2)  # 2 is for SW_MINIMIZE

# Show the form
$form.ShowDialog() | Out-Null