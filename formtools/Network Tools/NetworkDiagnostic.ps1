# Title: Network Tools

<#
    .SYNOPSIS
    A GUI-based Network Diagnostic Tool for pinging, tracing routes, and performing NSLookup operations.

    .DESCRIPTION
    This PowerShell tool offers a user-friendly interface to:
    1. Ping IP addresses or domains with configurable count and timeout.
    2. Trace the route to a destination with an optional setting to not resolve addresses.
    3. Perform NSLookup operations on IP addresses or domains.
    4. Display network adapter details and public IP.
    5. Perform DNS cache flush.
    6. Save results to a text file.
    The tool also provides interactive feedback and adjusts its interface based on user choices.

    .AUTHOR
    Eric Thorup

    .COPYRIGHT
    Copyright (c) 2023 TEK Utah LLC. All rights reserved.
#>

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

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Network Diagnostic Tool'
$form.Size = New-Object System.Drawing.Size(625, 540)
$form.MinimumSize = New-Object System.Drawing.Size(625, 540)  # Set minimum size
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen  # Center the form on screen
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
$form.Add_Activated({
    $form.TopMost = $true
})

# Radio Button for Ping
$radioPing = New-Object System.Windows.Forms.RadioButton
$radioPing.Location = New-Object System.Drawing.Point 10, 10
$radioPing.Size = New-Object System.Drawing.Size 60, 20
$radioPing.Text = 'Ping'
$radioPing.Checked = $true
$form.Controls.Add($radioPing) | Out-Null

# Radio Button for Tracert
$radioTracert = New-Object System.Windows.Forms.RadioButton
$radioTracert.Location = New-Object System.Drawing.Point 80, 10
$radioTracert.Size = New-Object System.Drawing.Size 80, 20
$radioTracert.Text = 'Tracert'
$form.Controls.Add($radioTracert) | Out-Null

# Radio Button for NSLookup
$radioNSLookup = New-Object System.Windows.Forms.RadioButton
$radioNSLookup.Location = New-Object System.Drawing.Point(160, 10) # Adjusted initial position
$radioNSLookup.Size = New-Object System.Drawing.Size(90, 20)
$radioNSLookup.Text = 'NSLookup'
$form.Controls.Add($radioNSLookup) | Out-Null

# Label for IP Address
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point 10, 40
$label.Size = New-Object System.Drawing.Size 120, 20
$label.Text = 'Enter IP or Domain:'
$form.Controls.Add($label) | Out-Null

# TextBox for IP Address
$textBoxIP = New-Object System.Windows.Forms.TextBox
$textBoxIP.Location = New-Object System.Drawing.Point 130, 40
$textBoxIP.Size = New-Object System.Drawing.Size 200, 20
$form.Controls.Add($textBoxIP) | Out-Null

# Button to start ping
$buttonGo = New-Object System.Windows.Forms.Button
$buttonGo.Location = New-Object System.Drawing.Point 500, 40
$buttonGo.Size = New-Object System.Drawing.Size 40, 20
$buttonGo.Text = 'Go'
$form.Controls.Add($buttonGo) | Out-Null

# Button to stop ping
$buttonStop = New-Object System.Windows.Forms.Button
$buttonStop.Location = New-Object System.Drawing.Point 550, 40
$buttonStop.Size = New-Object System.Drawing.Size 40, 20
$buttonStop.Text = 'Stop'
$form.Controls.Add($buttonStop) | Out-Null
$buttonStop.Enabled = $false

# Checkbox for saving results
$checkBoxSaveResults = New-Object System.Windows.Forms.CheckBox
$checkBoxSaveResults.Location = New-Object System.Drawing.Point 10, 280
$checkBoxSaveResults.Size = New-Object System.Drawing.Size 100, 20
$checkBoxSaveResults.Text = 'Save Results'
$form.Controls.Add($checkBoxSaveResults) | Out-Null

# Flush DNS Button
$buttonFlushDNS = New-Object System.Windows.Forms.Button
$buttonFlushDNS.Location = New-Object System.Drawing.Point ($checkBoxSaveResults.Right + 10), $checkBoxSaveResults.Location.Y
$buttonFlushDNS.Size = New-Object System.Drawing.Size(90, 20)
$buttonFlushDNS.Text = 'Flush DNS'
$form.Controls.Add($buttonFlushDNS) | Out-Null

$buttonFlushDNS.Add_Click({
    try {
        # Run the ipconfig /flushdns command
        $result = Invoke-Expression -Command "ipconfig /flushdns"
        
        # Show a message box to the user
        [System.Windows.Forms.MessageBox]::Show('DNS cache flushed.', 'Success', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show('Failed to flush DNS cache.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Button to Fix Network Stack
$buttonFixNetworkStack = New-Object System.Windows.Forms.Button
$buttonFixNetworkStack.Location = New-Object System.Drawing.Point ($buttonFlushDNS.Right + 10), $buttonFlushDNS.Location.Y
$buttonFixNetworkStack.Size = New-Object System.Drawing.Size(110, 20)
$buttonFixNetworkStack.Text = 'Fix Network Stack'
$form.Controls.Add($buttonFixNetworkStack) | Out-Null

$buttonFixNetworkStack.Add_Click({
    try {
        # Run the commands to fix the network stack
        Invoke-Expression -Command "netsh int ip reset"
        Invoke-Expression -Command "netsh winsock reset"
        
        # Show a message box to the user
        [System.Windows.Forms.MessageBox]::Show('Network stack fixed.', 'Success', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show('Failed to fix network stack.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# TextBox for displaying results is replaced by a TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 70)
$tabControl.Size = New-Object System.Drawing.Size(480, 200)
$form.Controls.Add($tabControl) | Out-Null

# TabPage for displaying results
$tabPageResult = New-Object System.Windows.Forms.TabPage
$tabPageResult.Text = "Results"
$tabControl.Controls.Add($tabPageResult) | Out-Null

# TextBox inside the Results TabPage
$textBoxResult = New-Object System.Windows.Forms.TextBox
$textBoxResult.Dock = [System.Windows.Forms.DockStyle]::Fill
$textBoxResult.Multiline = $true
$textBoxResult.ScrollBars = 'Vertical'
$tabPageResult.Controls.Add($textBoxResult) | Out-Null

# TabPage for Network Scan
$tabPageScan = New-Object System.Windows.Forms.TabPage
$tabPageScan.Text = "Network Scan"
$tabControl.Controls.Add($tabPageScan) | Out-Null

# ListView inside the Network Scan TabPage
$listViewScan = New-Object System.Windows.Forms.ListView
$listViewScan.Dock = [System.Windows.Forms.DockStyle]::Fill
$listViewScan.View = [System.Windows.Forms.View]::Details
$listViewScan.Columns.Add("IP Address", 120) | Out-Null
$listViewScan.Columns.Add("Hostname", 120) | Out-Null
$listViewScan.Columns.Add("Status", 80) | Out-Null
$tabPageScan.Controls.Add($listViewScan) | Out-Null

# Label for Number of Pings
$labelPingCount = New-Object System.Windows.Forms.Label
$labelPingCount.Location = New-Object System.Drawing.Point(500, 70)
$labelPingCount.Size = New-Object System.Drawing.Size(80, 20)
$labelPingCount.Text = 'Ping Count:'
$form.Controls.Add($labelPingCount) | Out-Null

# Label for Ping Timeout
$labelPingTimeout = New-Object System.Windows.Forms.Label
$labelPingTimeout.Location = New-Object System.Drawing.Point(500, 120) # Adjust the position accordingly
$labelPingTimeout.Size = New-Object System.Drawing.Size(80, 20)
$labelPingTimeout.Text = 'Ping Timeout:'
$form.Controls.Add($labelPingTimeout) | Out-Null

# UpDown Control for Number of Pings
$pingCountUpDown = New-Object System.Windows.Forms.NumericUpDown
$pingCountUpDown.Location = New-Object System.Drawing.Point(500, 90)
$pingCountUpDown.Size = New-Object System.Drawing.Size(60, 20)  # Adjusted width here
$pingCountUpDown.Maximum = [decimal]::MaxValue
$pingCountUpDown.Minimum = 0
$pingCountUpDown.Value = 0
$form.Controls.Add($pingCountUpDown) | Out-Null

# UpDown Control for Ping Timeout
$pingTimeoutUpDown = New-Object System.Windows.Forms.NumericUpDown
$pingTimeoutUpDown.Location = New-Object System.Drawing.Point(500, 140)
$pingTimeoutUpDown.Size = New-Object System.Drawing.Size(60, 20)  # Adjusted width here
$pingTimeoutUpDown.Maximum = 5000
$pingTimeoutUpDown.Minimum = 1
$pingTimeoutUpDown.Value = 1000
$form.Controls.Add($pingTimeoutUpDown) | Out-Null

# Checkbox for Tracert -d option
$checkBoxTracertNoResolve = New-Object System.Windows.Forms.CheckBox
$checkBoxTracertNoResolve.Location = New-Object System.Drawing.Point(500, 70)
$checkBoxTracertNoResolve.Size = New-Object System.Drawing.Size(100, 40)
$checkBoxTracertNoResolve.Text = 'Do not resolve addresses'
$checkBoxTracertNoResolve.Visible = $false  # <-- Set this to false initially
$form.Controls.Add($checkBoxTracertNoResolve) | Out-Null

# Create Label for Start IP Address in Network Scan TabPage
$labelStartIP = New-Object System.Windows.Forms.Label
$labelStartIP.Location = New-Object System.Drawing.Point 10, 40  # Adjusted to be next to the main IP/Domain input
$labelStartIP.Size = New-Object System.Drawing.Size 60, 20
$labelStartIP.Text = 'Start IP:'
$labelStartIP.Visible = $false
$form.Controls.Add($labelStartIP) | Out-Null

# Create TextBox for Start IP Address in Network Scan TabPage
$textBoxStartIP = New-Object System.Windows.Forms.TextBox
$textBoxStartIP.Location = New-Object System.Drawing.Point ($labelStartIP.Right + 5), 40  # Positioned right next to the label
$textBoxStartIP.Size = New-Object System.Drawing.Size 140, 20
$textBoxStartIP.Visible = $false
$form.Controls.Add($textBoxStartIP) | Out-Null

# Create Label for End IP Address in Network Scan TabPage
$labelEndIP = New-Object System.Windows.Forms.Label
$labelEndIP.Location = New-Object System.Drawing.Point ($textBoxStartIP.Right + 10), 40  # Positioned right next to Start IP TextBox
$labelEndIP.Size = New-Object System.Drawing.Size 60, 20
$labelEndIP.Text = 'End IP:'
$labelEndIP.Visible = $false
$form.Controls.Add($labelEndIP) | Out-Null

# Create TextBox for End IP Address in Network Scan TabPage
$textBoxEndIP = New-Object System.Windows.Forms.TextBox
$textBoxEndIP.Location = New-Object System.Drawing.Point ($labelEndIP.Right + 5), 40  # Positioned right next to the label
$textBoxEndIP.Size = New-Object System.Drawing.Size 140, 20
$textBoxEndIP.Visible = $false
$form.Controls.Add($textBoxEndIP) | Out-Null

# Create button for selecting IP range
$buttonSelectRange = New-Object System.Windows.Forms.Button
$buttonSelectRange.Location = New-Object System.Drawing.Point ($textBoxEndIP.Right + 5), 40  # Positioned right next to End IP TextBox
$buttonSelectRange.Size = New-Object System.Drawing.Size(50, 20)
$buttonSelectRange.Text = '<-->'
$buttonSelectRange.Visible = $false
$form.Controls.Add($buttonSelectRange) | Out-Null

$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabPageScan) {
        $label.Visible = $false
        $textBoxIP.Visible = $false
        
        $labelStartIP.Visible = $true
        $textBoxStartIP.Visible = $true
        $labelEndIP.Visible = $true
        $textBoxEndIP.Visible = $true
		$buttonSelectRange.Visible = $true
    } else {
        $label.Visible = $true
        $textBoxIP.Visible = $true
        
        $labelStartIP.Visible = $false
        $textBoxStartIP.Visible = $false
        $labelEndIP.Visible = $false
        $textBoxEndIP.Visible = $false
		$buttonSelectRange.Visible = $false
    }
})

$radioPing.Add_CheckedChanged({
    $labelPingCount.Visible = $true
    $pingCountUpDown.Visible = $true
    $labelPingTimeout.Visible = $true
    $pingTimeoutUpDown.Visible = $true
    $checkBoxTracertNoResolve.Visible = $false
})

$radioTracert.Add_CheckedChanged({
    $labelPingCount.Visible = $false
    $pingCountUpDown.Visible = $false
    $labelPingTimeout.Visible = $false
    $pingTimeoutUpDown.Visible = $false
    $checkBoxTracertNoResolve.Visible = $true
})

$radioPing.Add_CheckedChanged({
    $labelPingCount.Visible = $true
    $pingCountUpDown.Visible = $true
    $labelPingTimeout.Visible = $true
    $pingTimeoutUpDown.Visible = $true
})

$radioTracert.Add_CheckedChanged({
    $labelPingCount.Visible = $false
    $pingCountUpDown.Visible = $false
    $labelPingTimeout.Visible = $false
    $pingTimeoutUpDown.Visible = $false
})

$radioNSLookup.Add_CheckedChanged({
    $labelPingCount.Visible = $false
    $pingCountUpDown.Visible = $false
    $labelPingTimeout.Visible = $false
    $pingTimeoutUpDown.Visible = $false
    $checkBoxTracertNoResolve.Visible = $false
})


# Retrieve the IP information for all network adapters using WMI
$adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object {
    $_.IPEnabled -eq $true -and $_.Description -notmatch "Loopback|vEthernet"
}

# Initial positions for the dynamically created text boxes
$initialX = 10
$initialY = $checkBoxSaveResults.Bottom + 10
$currentX = $initialX
$currentY = $initialY
$textBoxes = @() # To store references to the created text boxes

function Get-ActiveNetworks {
    $activeConnections = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object {
        $_.IPEnabled -eq $true -and $_.Description -notmatch "Loopback|vEthernet" -and $_.IPAddress
    }
    
    function ConvertIPToUInt32 ($ip) {
        $bytes = $ip.Split('.') | ForEach-Object { [byte]$_ }
        [BitConverter]::ToUInt32($bytes[0..3], 0)
    }

    $connectionDetails = @()
    foreach ($conn in $activeConnections) {
        $ipAddressInt = ConvertIPToUInt32 $conn.IPAddress[0]
        $subnetMaskInt = ConvertIPToUInt32 $conn.IPSubnet[0]

        $ipSubnetInt = $ipAddressInt -band $subnetMaskInt
        $ipSubnet = [IPAddress]$ipSubnetInt

        $prefixLength = ([IPAddress]$conn.IPSubnet[0]).GetAddressBytes() | ForEach-Object { [Convert]::ToString($_, 2) } | ForEach-Object { $_.ToCharArray() } | Where-Object { $_ -eq '1' } | Measure-Object | Select-Object -ExpandProperty Count

        $connectionDetails += @{
            'IPSubnet' = "$($ipSubnet.ToString())/$prefixLength"
            'IPAddress' = $conn.IPAddress[0]
        }
    }

    return $connectionDetails
}

function CreateClickHandler {
    param($ipAddress)

    return {
        $global:startIP = $ipAddress
        $textBoxStartIP.Text = $global:startIP
        $textBoxEndIP.Text = $global:startIP
        $dialog.Close()
    }
}


function Show-SubnetDialog {
    $subnets = Get-ActiveNetworks
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Select IP Address'
    
    # Adjust dialog size
    $dialog.Size = New-Object System.Drawing.Size(240, ($subnets.Count * 65 + 70))
    $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    
    # Adjust label size
    $label.Size = New-Object System.Drawing.Size(200, 40)
    $label.Text = 'Which IP address do you want to select?'
    $dialog.Controls.Add($label) | Out-Null

    $y = 60
    $buttonSpacing = 40
    foreach ($subnet in $subnets) {
        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(10, $y)
        
        # Adjust button size
        $button.Size = New-Object System.Drawing.Size(200, 30)
        $button.Text = $subnet.IPAddress  # Set button text to IP address
        $button.Tag = $subnet.IPAddress   # Set the Tag property to IP address
        
        $button.Add_Click({
            param($sender, $e)
            $clickedButton = $sender
            $global:startIP = $clickedButton.Tag
            $textBoxStartIP.Text = $global:startIP
            $textBoxEndIP.Text = $global:startIP
            $dialog.Close()
        })
        
        $dialog.Controls.Add($button) | Out-Null
        
        # Adjust vertical spacing for next button
        $y += $buttonSpacing
    }
    $form.TopMost = $false
    $dialog.Owner = $form
    $dialog.ShowDialog()
    $form.TopMost = $true
}

$buttonSelectRange.Add_Click({
    Show-SubnetDialog
})

foreach ($adapter in $adapters) {
	# Concatenate DNS servers with newline
    $dnsServers = $adapter.DNSServerSearchOrder -join "`r`n"

    # Create a new text box for the adapter
    $textBoxAdapter = New-Object System.Windows.Forms.TextBox
    $textBoxAdapter.Location = New-Object System.Drawing.Point $currentX, $currentY
    $textBoxAdapter.Size = New-Object System.Drawing.Size ($form.ClientSize.Width / 3 - 20), 150
    $textBoxAdapter.Multiline = $true
    $textBoxAdapter.ScrollBars = 'Vertical'
    $textBoxAdapter.Text = "Description: $($adapter.Description)`r`nIP Address: $($adapter.IPAddress[0])`r`nSubnet Mask: $($adapter.IPSubnet[0])`r`nDefault Gateway: $($adapter.DefaultIPGateway)`r`nDNS Servers: $dnsServers"
    $form.Controls.Add($textBoxAdapter) | Out-Null
    $textBoxes += $textBoxAdapter

	# Adjust the X and Y positions for the next text box
	if (($currentX + $textBoxAdapter.Width * 2 + 10) -le $form.ClientSize.Width) {
		$currentX += $textBoxAdapter.Width + 10
	} else {
		$currentX = $initialX
		$currentY += $textBoxAdapter.Height + 10
	}
}

# Adjust text box sizes and positions when the form is resized or layout changes
$adjustTextBoxes = {
    $x = $initialX
    $y = $checkBoxSaveResults.Bottom + 10

    # Adjust positions for network info boxes
    foreach ($textBox in $textBoxes) {
        $textBox.Width = ($form.ClientSize.Width / 3 - 20)
        $textBox.Location = New-Object System.Drawing.Point $x, $y

        # Check if next position (with another textbox width) would exceed the form width
        if (($x + $textBox.Width * 2 + 10) -le $form.ClientSize.Width) {
            $x += $textBox.Width + 10
        } else {
            $x = $initialX
            $y += $textBox.Height + 10
        }
    }

    # Adjust the position for Public IP controls based on the last network info box
    $labelPublicIP.Location = New-Object System.Drawing.Point 10, ($textBoxes[-1].Bottom + 10)
    $textBoxPublicIP.Location = New-Object System.Drawing.Point 130, ($textBoxes[-1].Bottom + 10)
}

# Label for Public IP Address
$labelPublicIP = New-Object System.Windows.Forms.Label
$labelPublicIP.Location = New-Object System.Drawing.Point 10, ($textBoxes[-1].Bottom + 10)  # Place it below the last network info box
$labelPublicIP.Size = New-Object System.Drawing.Size 120, 20
$labelPublicIP.Text = 'Public IP Address:'
$form.Controls.Add($labelPublicIP) | Out-Null

# TextBox for Public IP Address
$textBoxPublicIP = New-Object System.Windows.Forms.TextBox
$textBoxPublicIP.Location = New-Object System.Drawing.Point 130, ($textBoxes[-1].Bottom + 10)
$textBoxPublicIP.Size = New-Object System.Drawing.Size 200, 20
$textBoxPublicIP.ReadOnly = $true  # Make it read-only
$form.Controls.Add($textBoxPublicIP) | Out-Null

# Fetch and display the public IP
try {
    $ipInfo = Invoke-RestMethod -Uri 'http://ipinfo.io/json'
    $textBoxPublicIP.Text = $ipInfo.ip
} catch {
    $textBoxPublicIP.Text = "Failed to retrieve public IP"
}

$form.Add_Resize({
    # Adjust the ping count, ping timeout, and associated labels
    $labelPingCount.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - $labelPingCount.Width - 20), $labelPingCount.Location.Y
    $pingCountUpDown.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - $pingCountUpDown.Width - 20), $pingCountUpDown.Location.Y
    $labelPingTimeout.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - $labelPingTimeout.Width - 20), $labelPingTimeout.Location.Y
    $pingTimeoutUpDown.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - $pingTimeoutUpDown.Width - 20), $pingTimeoutUpDown.Location.Y
    
	$buttonGo.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - $buttonGo.Width - $buttonStop.Width - 10), 40)
	$buttonStop.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - $buttonStop.Width - 5), 40)

    # Adjust the 'Do not resolve addresses' checkbox
    $checkBoxTracertNoResolve.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - $checkBoxTracertNoResolve.Width - 20), $checkBoxTracertNoResolve.Location.Y

    # Adjust the vertical position of the 'Flush DNS' button based on the 'Save Results' checkbox
    $buttonFlushDNS.Location = New-Object System.Drawing.Point ($checkBoxSaveResults.Right + 10), $checkBoxSaveResults.Location.Y

    #  Adjust the NSLookup radio button's X position
    $radioNSLookup.Location = New-Object System.Drawing.Point($radioTracert.Location.X + $radioTracert.Width + 10), 10

    # Adjust the 'Save Results' checkbox to be below the output textbox
    $checkBoxSaveResults.Location = New-Object System.Drawing.Point -ArgumentList $checkBoxSaveResults.Location.X, ($tabControl.Bottom + 10)

    # Adjust the vertical position of the 'Flush DNS' button based on the 'Save Results' checkbox
    $buttonFlushDNS.Location = New-Object System.Drawing.Point ($checkBoxSaveResults.Right + 10), $checkBoxSaveResults.Location.Y

    # Adjust the network adapter information text boxes to be below the 'Flush DNS' button
    $x = $initialX
    $y = $buttonFlushDNS.Bottom + 10
    foreach ($textBox in $textBoxes) {
        $textBox.Location = New-Object System.Drawing.Point $x, $y

        if (($x + $textBox.Width * 2 + 10) -le $form.ClientSize.Width) {
            $x += $textBox.Width + 10
        } else {
            $x = $initialX
            $y += $textBox.Height + 10
        }
    }

    # Adjust the position for Public IP controls based on the last network info box
    $labelPublicIP.Location = New-Object System.Drawing.Point 10, ($textBoxes[-1].Bottom + 10)
    $textBoxPublicIP.Location = New-Object System.Drawing.Point 130, ($textBoxes[-1].Bottom + 10)
})

# Define a global variable to indicate whether the trace complete dialog has been shown
$traceCompleteShown = $false

$form.Add_Resize($adjustTextBoxes)
$form.Add_Layout($adjustTextBoxes)

# Adjust the anchor properties for auto-resizing
$textBoxIP.Anchor = 'Top,Left,Right'
$textBoxStartIP.Anchor = 'Top,Left,Right'
$textBoxEndIP.Anchor = 'Top,Left,Right'
$buttonGo.Anchor = 'Top,Right'
$buttonStop.Anchor = 'Top,Right'
$tabControl.Anchor = 'Top,Left,Right,Bottom'

# Define the main timer for ping and trace RT actions
$mainTimer = New-Object System.Windows.Forms.Timer
$mainTimer.Interval = [int]$pingTimeoutUpDown.Value + 50

# Define a separate timer for trace RT action
$tracertTimer = $null

$nslookupCompletionTimer = New-Object System.Windows.Forms.Timer
$nslookupCompletionTimer.Interval = 4000  # 4 seconds

$nslookupCompletionTimer.Add_Tick({
    $buttonGo.Enabled = $true
    $nslookupCompletionTimer.Stop()
})

$nslookupAction = {
    $textBoxResult.Clear()
    
    try {
        $nslookupResults = Resolve-DnsName -Name $global:ip
        
        foreach ($result in $nslookupResults) {
            $textBoxResult.AppendText(($result | Out-String) + "`r`n")
        }
        
        Save-Results  # Save the results if needed

        # Start the timer to re-enable the Go button after 4 seconds
        $nslookupCompletionTimer.Start()

    } catch {
        $textBoxResult.AppendText("NSLookup failed: $_`r`n")
        $buttonGo.Enabled = $true
		$buttonStop.Enabled = $false

    }
}

Function Save-Results {
    if ($checkBoxSaveResults.Checked) {
        $timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $defaultFilePath = ""

        if ($radioPing.Checked) {
            $defaultFilePath = "Logs\pingout_$timestamp.txt"
            $global:pingResults -join "`r`n" | Out-File -FilePath $defaultFilePath
        } elseif ($radioTracert.Checked -and (Test-Path "Logs\traceout.txt")) {
            # Delay for 1 second to allow the file handle to release
            Start-Sleep -Seconds 1
            
            # Rename the trace route output file
            $defaultFilePath = "traceout_$timestamp.txt"
            Rename-Item -Path "Logs\traceout.txt" -NewName $defaultFilePath
			$defaultFilePath = "Logs\traceout_$timestamp.txt"
        } elseif ($radioNSLookup.Checked) {
            $defaultFilePath = "Logs\nslookup_$timestamp.txt"
            [System.IO.File]::AppendAllText($defaultFilePath, $textBoxResult.Text)
        }
		Start-Sleep -Seconds 1
        # Prompt user to save a copy of the file
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
        $saveFileDialog.Title = 'Save a Copy of Results'
        
        $result = $saveFileDialog.ShowDialog()

        if ($result -eq 'OK') {
            Copy-Item -Path $defaultFilePath -Destination $saveFileDialog.FileName
        }
    } else {
        if (Test-Path "Logs\traceout.txt") {
            # Delay for 1 second to allow the file handle to release
            Start-Sleep -Seconds 1
            
            # Delete the temporary output file
            Remove-Item -Path "Logs\traceout.txt" -ErrorAction SilentlyContinue
        }
    }
}

# Define a global variable to store ping results
$global:pingResults = @()

$ping = New-Object System.Net.NetworkInformation.Ping

$global:pingAdded = $false

$global:pingRemaining = 0

$pingAction = {
    $timeout = [int]$pingTimeoutUpDown.Value
    if ($global:pingRemaining -eq 0) {
        $global:pingRemaining = [int]$pingCountUpDown.Value
    }
    try {
        $reply = $ping.Send($global:ip, $timeout)
        if ($reply.Status -eq 'Success') {
            $resultText = "Reply from $($reply.Address): bytes=$($reply.Buffer.Length) time=$($reply.RoundtripTime)ms TTL=$($reply.Options.Ttl)`r`n"
            $textBoxResult.AppendText($resultText)
            # Store result in the global variable
            $global:pingResults += $resultText
        } else {
            $resultText = "Request timed out.`r`n"
            $textBoxResult.AppendText($resultText)
            # Store result in the global variable
            $global:pingResults += $resultText
        }
    } catch {
        if ($_.Exception.InnerException -ne $null) {
            $resultText = "Ping failed: " + $_.Exception.InnerException.Message + "`r`n"
        } else {
            $resultText = "Ping failed: $_.Exception.Message`r`n"
        }
        $textBoxResult.AppendText($resultText)
        # Store result in the global variable
        $global:pingResults += $resultText
    }
    $textBoxResult.SelectionStart = $textBoxResult.Text.Length
    $textBoxResult.ScrollToCaret()

	# Decrement the ping remaining count if it's greater than 0
	if ($global:pingRemaining -gt 0) {
		$global:pingRemaining--
		# If no pings remaining, stop the timer, save results and re-enable the Go button
		if ($global:pingRemaining -eq 0) {
			$mainTimer.Stop()
			Save-Results
			$buttonGo.Enabled = $true
			$buttonStop.Enabled = $false

		}
	}
}

$tracertAction = {
    $textBoxResult.Clear()

    # Initialize flag to false
    $traceCompleteShown = $false
    
    # Check the checkbox state and adjust the command accordingly
    $tracertCmdArgs = if ($checkBoxTracertNoResolve.Checked) {
        "-d $($textBoxIp.Text)"
    } else {
        $textBoxIp.Text
    }

    $process = Start-Process tracert -ArgumentList $tracertCmdArgs -NoNewWindow -PassThru -RedirectStandardOutput "Logs\traceout.txt"
    
    $tracertTimer = New-Object System.Windows.Forms.Timer
    $tracertTimer.Interval = 1000

    $tracertTimer.Add_Tick({
        # Check if the traceout.txt file exists in the "Logs" directory
        if (Test-Path "Logs\traceout.txt") {
            $content = Get-Content "Logs\traceout.txt" -Raw
            $splitContent = $content -split "\r\n" | Where-Object {$_ -notmatch "^\s*\d+\s+ms\s*$"}
            $textBoxResult.Text = $splitContent -join "`r`n"
        }

        # Check if the content contains "Trace complete."
        if ($content -like "*Trace complete.*" -and !$traceCompleteShown) {
            $traceCompleteShown = $true  # Set flag to true

			# Call Save-Results to handle saving or deleting the file
			Save-Results

			# Check if $tracertTimer is not null before calling Stop and Dispose
			if ($tracertTimer) {
				$tracertTimer.Stop()
				$tracertTimer.Dispose()
				$tracertTimer = $null
			}

			# Re-enable the Go button
			$buttonGo.Enabled = $true
			$buttonStop.Enabled = $false

		}
    })

    $tracertTimer.Start()
}

# actions for Go button
$buttonGo.Add_Click({
    # First, check which tab is selected
    if ($tabControl.SelectedTab -eq $tabPageResult) {
        # Remove the ping and trace RT actions from the main timer's Tick event
        $mainTimer.Remove_Tick($pingAction)
        $mainTimer.Remove_Tick($tracertAction)
	    $buttonStop.Enabled = $true

        $global:ip = $textBoxIP.Text.Trim()
        $textBoxResult.Clear()

        if (-not [string]::IsNullOrEmpty($global:ip)) {
            $buttonGo.Enabled = $false  # Disable the Go button

            # Clear previous results
            $global:pingResults = @()

            # Reset the global ping remaining counter
            $global:pingRemaining = [int]$pingCountUpDown.Value

            if ($radioPing.Checked) {
                $mainTimer.Interval = [int]$pingTimeoutUpDown.Value + 50

                # Add the ping action to the main timer's Tick event
                $mainTimer.Add_Tick($pingAction)
                $mainTimer.Start()  # Start the main timer only for the Ping action
            } elseif ($radioTracert.Checked) {
                & $tracertAction  # Execute the Trace RT action directly
            } elseif ($radioNSLookup.Checked) {
                & $nslookupAction  # Execute the NSLookup action directly
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show('Please enter a valid IP address or domain name.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } elseif ($tabControl.SelectedTab -eq $tabPageScan) {
        # Logic for when the Network Scan tab is selected goes here.
        # For now, it's left blank. You can add functionality specific to this tab.
    }
})

# Stop action for Stop button
$buttonStop.Add_Click({
    $mainTimer.Stop()  # Stop the main timer

    # Stop and dispose of the tracert timer if it's running
    if ($tracertTimer -ne $null) {
        $tracertTimer.Stop()
        $tracertTimer.Dispose()
        $tracertTimer = $null
    }

    # Attempt to stop the Tracert process gracefully
    Stop-Process -Name tracert -ErrorAction SilentlyContinue

    Save-Results  # Save Results
    $buttonGo.Enabled = $true  # Re-enable the Go button
	$buttonStop.Enabled = $false

})

# Minimize the PowerShell console window
$consoleWindowHandle = [ConsoleWindow]::GetConsoleWindow()
[ConsoleWindow]::ShowWindow($consoleWindowHandle, 2) | Out-Null  # 2 is for SW_MINIMIZE

$form.ShowDialog() | Out-Null