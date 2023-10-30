<#
    .SYNOPSIS
    A PowerShell script interface to run PowerShell scripts, install software, and access web shortcuts and access to a system info script.

    .DESCRIPTION
    This GUI-based PowerShell tool allows users to:
    1. Run PowerShell scripts with optional command line arguments.
    2. Install software by executing .exe or .msi files  with optional command line arguments.
    3. Access web shortcuts by opening .url files.
    The interface supports drag-and-drop functionality, and items can be executed with or without administrative privileges.

    .AUTHOR
    Eric Thorup

    .COPYRIGHT
    Copyright (c) 2023 TEK Utah LLC. All rights reserved.
#>

# Get the current script or exe's directory
$currentDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

Add-Type -AssemblyName System.Windows.Forms
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
$form.Text = "PowerShell Script Runner"
$form.Size = New-Object System.Drawing.Size(500,400)
$form.MinimumSize = New-Object System.Drawing.Size(500,400)
$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Add_Activated({
    $form.TopMost = $true
})

# Create the TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabControl.Height = $form.Height - 150  # Adjusting the height to ensure it doesn't cover the controls at the bottom
$tabControl.Location = New-Object System.Drawing.Point(10,30)
$tabControl.Size = New-Object System.Drawing.Size(460,160)
$form.Controls.Add($tabControl) | Out-Null

# Create the PowerShell Scripts tab and add to TabControl
$psScriptsTab = New-Object System.Windows.Forms.TabPage
$psScriptsTab.Text = "PowerShell Scripts"
$tabControl.Controls.Add($psScriptsTab) | Out-Null

# Create the Install Software tab and add to TabControl
$installSoftwareTab = New-Object System.Windows.Forms.TabPage
$installSoftwareTab.Text = "Install Software"
$tabControl.Controls.Add($installSoftwareTab) | Out-Null

# Create the WebShortcuts tab and add to TabControl
$webShortcutsTab = New-Object System.Windows.Forms.TabPage
$webShortcutsTab.Text = "WebShortcuts"
$tabControl.Controls.Add($webShortcutsTab) | Out-Null

# Create the list box for PowerShell Scripts and add to the PowerShell Scripts tab
$psScriptsListBox = New-Object System.Windows.Forms.ListBox
$psScriptsListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$psScriptsListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$psScriptsListBox.AllowDrop = $true
$psScriptsListBox.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

function RefreshPsScriptsList {
    $psScriptsListBox.Items.Clear() # Clear current items

    Get-ChildItem -Path $scriptPath -Filter "*.ps1" | ForEach-Object {
        $lines = Get-Content $_.FullName -TotalCount 2

        if ($lines.Count -ge 1 -and $lines[0] -match "#Title: ") {
            $title = ($lines[0] -split "#Title: ")[1].Trim()
        } else {
            $title = $_.BaseName
        }

        $psScriptsListBox.Items.Add($title) | Out-Null
        $scriptMapping[$title] = $_.FullName
    }
}

$psScriptsListBox.Add_DragDrop({
    $e = $_
    $file = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)[0]
    
    if ($file -match "\.ps1$") {
        Copy-Item -Path $file -Destination $scriptPath
        RefreshPSScriptsList
    } else {
        $form.TopMost = $true
        [System.Windows.Forms.MessageBox]::Show($form, "Only .ps1 files are allowed!", "Invalid File Type", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.TopMost = $false
    }
})
$psScriptsTab.Controls.Add($psScriptsListBox) | Out-Null

# Create the list box for Install Software and add to the Install Software tab
$softwareListBox = New-Object System.Windows.Forms.ListBox
$softwareListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$softwareListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$softwareListBox.AllowDrop = $true
$softwareListBox.Add_DragEnter({
    $e = $_
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [Windows.Forms.DragDropEffects]::Copy
    } else {
        $e.Effect = [Windows.Forms.DragDropEffects]::None
    }
})

$softwareListBox.Add_DragDrop({
    $e = $_
    $file = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)[0]
    
    if ($file -match "\.(exe|msi)$") {
        Copy-Item -Path $file -Destination $softwarePath
        RefreshSoftwareList
    } else {
        $form.TopMost = $true
        [System.Windows.Forms.MessageBox]::Show($form, "Only .exe and .msi files are allowed!", "Invalid File Type", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.TopMost = $false
    }
})
$installSoftwareTab.Controls.Add($softwareListBox) | Out-Null

function RefreshSoftwareList {
    $softwareListBox.Items.Clear() # Clear current items

    Get-ChildItem -Path $softwarePath -Include @("*.exe","*.msi") -Recurse | ForEach-Object {
        $softwareListBox.Items.Add($_.Name) | Out-Null
    }
}

# Create the list box for WebShortcuts and add to the WebShortcuts tab
$webShortcutsListBox = New-Object System.Windows.Forms.ListBox
$webShortcutsListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$webShortcutsListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$webShortcutsListBox.AllowDrop = $true
$webShortcutsListBox.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})
$webShortcutsListBox.Add_DragDrop({
    $e = $_
    $file = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)[0]
    
    if ($file -match "\.url$") {
        Copy-Item -Path $file -Destination $webShortcutsPath
        RefreshWebShortcutsList
    } else {
        $form.TopMost = $true
        [System.Windows.Forms.MessageBox]::Show($form, "Only .url files are allowed!", "Invalid File Type", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.TopMost = $false
    }
})
$webShortcutsTab.Controls.Add($webShortcutsListBox) | Out-Null

function RefreshWebShortcutsList {
    $webShortcutsListBox.Items.Clear() # Clear current items

    Get-ChildItem -Path $webShortcutsPath -Filter "*.url" | ForEach-Object {
        $webShortcutsListBox.Items.Add($_.BaseName) | Out-Null
    }
}

function AdjustControlPositions {
    # Adjust the position of the description box
    $textBox.Top = $tabControl.Bottom + 5

    # Adjust the position of the command line arguments label
    $cmdArgLabel.Top = $textBox.Bottom + 5

    # Adjust the position of the command line arguments text box
    $cmdArgTextBox.Top = $cmdArgLabel.Bottom + 5
}

# Create a hashtable to store the mapping between script titles and their paths
$scriptMapping = @{}

# Populate the psScriptsListBox with script titles
$scriptPath = Join-Path -Path $currentDir -ChildPath "scripts"
Get-ChildItem -Path $scriptPath -Filter "*.ps1" | ForEach-Object {
    $lines = Get-Content $_.FullName -TotalCount 2
    
    if ($lines.Count -ge 1 -and $lines[0] -match "#Title: ") {
        $title = ($lines[0] -split "#Title: ")[1].Trim()
    } else {
        $title = $_.BaseName
    }
    
    $psScriptsListBox.Items.Add($title) | Out-Null
    $scriptMapping[$title] = $_.FullName
}

# Populate the softwareListBox with exe and msi files from the software subfolder
$softwarePath = Join-Path -Path $currentDir -ChildPath "software"
Get-ChildItem -Path $softwarePath -Include @("*.exe","*.msi") -Recurse | ForEach-Object {
    $softwareListBox.Items.Add($_.Name) | Out-Null
}

# Populate the webShortcutsListBox with internet shortcuts
$webShortcutsPath = Join-Path -Path $currentDir -ChildPath "WebShortcuts"
Get-ChildItem -Path $webShortcutsPath -Filter "*.url" | ForEach-Object {
    $webShortcutsListBox.Items.Add($_.BaseName) | Out-Null
}

# Event for psScriptsListBox item selection
$psScriptsListBox.Add_SelectedIndexChanged({
    if ($psScriptsListBox.SelectedIndex -eq -1) { return } # Add this line

    $selectedTitle = $psScriptsListBox.SelectedItem
    $scriptFullPath = $scriptMapping[$selectedTitle] # Using the mapping to get the full path

    if (Test-Path $scriptFullPath) {
        $descriptionLines = Get-Content $scriptFullPath -TotalCount 2
        
        if ($descriptionLines.Count -ge 2 -and $descriptionLines[1] -match "#Description: ") {
            $description = $descriptionLines[1] -replace '#Description: ', ''
        } else {
            $description = ""
        }
        
        $textBox.Text = $description.Trim()
    }
})

# Create the text box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,200)
$textBox.Size = New-Object System.Drawing.Size(460,50)
$textBox.Multiline = $true
$textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox) | Out-Null

# Create a label for the command line arguments text box
$cmdArgLabel = New-Object System.Windows.Forms.Label
$cmdArgLabel.Location = New-Object System.Drawing.Point(10,260)
$cmdArgLabel.Size = New-Object System.Drawing.Size(200,20)
$cmdArgLabel.Text = "Command Line Arguments:"
$form.Controls.Add($cmdArgLabel) | Out-Null

# Create the text box for command line arguments
$cmdArgTextBox = New-Object System.Windows.Forms.TextBox
$cmdArgTextBox.Location = New-Object System.Drawing.Point(10,280)
$cmdArgTextBox.Size = New-Object System.Drawing.Size(460,20)
$cmdArgTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($cmdArgTextBox) | Out-Null

# Create Run as Administrator checkbox
$runAsAdminCheckbox = New-Object System.Windows.Forms.CheckBox
$runAsAdminCheckbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 7)
$runAsAdminCheckbox.Text = "Run as Administrator"
$runAsAdminCheckbox.Location = New-Object System.Drawing.Point(10, 320) # Adjusted based on form size
$runAsAdminCheckbox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($runAsAdminCheckbox) | Out-Null

# Create the run button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Location = New-Object System.Drawing.Point(370, 315) # Adjusted based on form size
$runButton.Size = New-Object System.Drawing.Size(100,30)
$runButton.Text = "Run"
$runButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($runButton) | Out-Null

$runButton.Add_Click({
    try {
        if ($tabControl.SelectedTab -eq $psScriptsTab) {
            # Check if an item is selected in the $psScriptsListBox
            if (-not $psScriptsListBox.SelectedItem) {
                [System.Windows.Forms.MessageBox]::Show("Please select a script from the list before clicking Run.", "Selection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            $selectedScriptTitle = $psScriptsListBox.SelectedItem
            $scriptFullPath = $scriptMapping[$selectedScriptTitle]
            $scriptContent = (Get-Content $scriptFullPath -Raw)
            $fullCommand = "$scriptContent $($cmdArgTextBox.Text)"

            if ($scriptFullPath -and (Test-Path $scriptFullPath)) {
                if ($runAsAdminCheckbox.Checked) {
                    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFullPath`" $cmdArgTextBox.Text"
                    Start-Process powershell.exe $arguments -Verb RunAs
                } else {
                    Invoke-Expression -Command $fullCommand
                }
            }
        } elseif ($tabControl.SelectedTab -eq $installSoftwareTab) {
            # Check if any items are selected in the $softwareListBox
            if ($softwareListBox.SelectedItems.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select software from the list before clicking Run.", "Selection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            foreach ($selectedSoftware in $softwareListBox.SelectedItems) {
                $softwareFullPath = Join-Path $softwarePath $selectedSoftware
                $softwareArguments = if ($cmdArgTextBox.Text) { $cmdArgTextBox.Text } else { $null }

                if (Test-Path $softwareFullPath) {
                    if ($runAsAdminCheckbox.Checked) {
                        if ([string]::IsNullOrEmpty($softwareArguments)) {
                            Start-Process -FilePath $softwareFullPath -Verb RunAs -Wait
                        } else {
                            Start-Process -FilePath $softwareFullPath -ArgumentList $softwareArguments -Verb RunAs -Wait
                        }
                    } else {
                        if ([string]::IsNullOrEmpty($softwareArguments)) {
                            Start-Process -FilePath $softwareFullPath -Wait
                        } else {
                            Start-Process -FilePath $softwareFullPath -ArgumentList $softwareArguments -Wait
                        }
                    }
                }
            }
        } elseif ($tabControl.SelectedTab -eq $webShortcutsTab) {
            # Check if an item is selected in the $webShortcutsListBox
            if (-not $webShortcutsListBox.SelectedItem) {
                [System.Windows.Forms.MessageBox]::Show("Please select a web shortcut from the list before clicking Run.", "Selection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }

            $selectedShortcut = $webShortcutsListBox.SelectedItem
            $shortcutFullPath = Join-Path $webShortcutsPath "$selectedShortcut.url"
            if (Test-Path $shortcutFullPath) {
                Start-Process -FilePath $shortcutFullPath
            } else {
                Write-Host "$shortcutFullPath does not exist!"
            }
        }
    } catch {
        if ($_.Exception.Message -match "cannot be stopped" -or $_.Exception.Message -match "cannot be started") {
            [System.Windows.Forms.MessageBox]::Show("This action requires administrative privileges. Please check the 'Run as Administrator' checkbox and try again.", "Admin Privileges Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } else {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$form.Controls.Add($runButton) | Out-Null

# Create MenuStrip and add to form
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.MainMenuStrip = $menuStrip

# Create File menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$menuStrip.Items.Add($fileMenu) | Out-Null

# Exit menu item under File
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.Add_Click({ $form.Close() })
$fileMenu.DropDownItems.Add($exitMenuItem) | Out-Null

function RunToolWithAdmin {
    param (
        [Parameter(Mandatory = $true)]
        [string]$toolPath
    )

    if (Test-Path $toolPath) {
        switch -Regex ($toolPath) {
            '\.ps1$' {
                # PowerShell script
                $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$toolPath`""
                Start-Process powershell.exe $arguments -Verb RunAs
            }
            '\.exe$' {
                # Executable file
                Start-Process -FilePath $toolPath -Verb RunAs
            }
            '\.msi$' {
                # MSI installer
                Start-Process msiexec.exe -ArgumentList "/i `"$toolPath`"" -Verb RunAs
            }
            '\.bat$' {
                # Batch file
                Start-Process cmd.exe -ArgumentList "/c `"$toolPath`"" -Verb RunAs
            }
            '\.vbs$' {
                # VBScript
                Start-Process cscript.exe -ArgumentList "`"$toolPath`"" -Verb RunAs
            }
            default {
                Write-Host "Unsupported file type for $toolPath"
            }
        }
    }
}

# Create a function to load tools into a menu, including nested directories
function LoadToolsMenu {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ToolStripMenuItem]$ParentMenu,
        [Parameter(Mandatory = $true)]
        [string]$ToolDir
    )

    # Clear any existing tools
    $ParentMenu.DropDownItems.Clear()

    # Populate tools from the directory
    if (Test-Path $ToolDir) {
        $items = Get-ChildItem -Path $ToolDir

        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                # Create a new sub-menu for the directory
                $subMenu = New-Object System.Windows.Forms.ToolStripMenuItem
                $subMenu.Text = $item.Name
                $ParentMenu.DropDownItems.Add($subMenu) | Out-Null

                # Recursive call to populate the sub-menu
                LoadToolsMenu -ParentMenu $subMenu -ToolDir $item.FullName
            } else {
                if ($item.Name -match "\.(ps1|exe|msi|vbs|bat)$") {
                    $toolMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
                    $toolMenuItem.Text = $item.BaseName
                    $toolMenuItem.Tag = $item.FullName  # Use the Tag property to store the full path
                    $toolMenuItem.Add_Click({
                        param($sender, $e)
                        $toolPath = $sender.Tag  # Retrieve the full path from the Tag property
                        RunToolWithAdmin -toolPath $toolPath
                    })
                    $ParentMenu.DropDownItems.Add($toolMenuItem) | Out-Null
                }
            }
        }
    }
}

# Create Tools menu
$toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolsMenu.Text = "Tools"
$menuStrip.Items.Add($toolsMenu) | Out-Null

# Load tools into the Tools menu
$formtoolsPath = Join-Path -Path $currentDir -ChildPath "formtools"
LoadToolsMenu -ParentMenu $toolsMenu -ToolDir $formtoolsPath

# Add MenuStrip to form
$form.Controls.Add($menuStrip) | Out-Null

# Event for tabControl tab selection
$tabControl.Add_SelectedIndexChanged({
    # Clear the content of description and command line argument text boxes
    $textBox.Text = ""
    $cmdArgTextBox.Text = ""
    # Uncheck the 'Run as Administrator' checkbox
    $runAsAdminCheckbox.Checked = $false
})

# add a right click context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$runMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$runMenuItem.Text = "Run"
$runMenuItem.Add_Click({
    foreach ($selectedScriptTitle in $psScriptsListBox.SelectedItems) {
        $scriptFullPath = $scriptMapping[$selectedScriptTitle]
        if ($scriptFullPath -and (Test-Path $scriptFullPath)) {
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFullPath`"" -Wait
        }
    }
})
$contextMenu.Items.Add($runMenuItem) | Out-Null

$runAsAdminScriptMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$runAsAdminScriptMenuItem.Text = "Run as Administrator"
$runAsAdminScriptMenuItem.Add_Click({
    # Code to run the PowerShell script with elevated privileges
    foreach ($selectedScriptTitle in $psScriptsListBox.SelectedItems) {
        $scriptFullPath = $scriptMapping[$selectedScriptTitle]
        if ($scriptFullPath -and (Test-Path $scriptFullPath)) {
            $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFullPath`" $cmdArgTextBox.Text"
            Start-Process powershell.exe $arguments -Verb RunAs -Wait
        }
    }
})
$contextMenu.Items.Add($runAsAdminScriptMenuItem) | Out-Null

$deleteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteMenuItem.Text = "Delete"
$deleteMenuItem.Add_Click({
    $selectedScriptTitle = $psScriptsListBox.SelectedItem
    $scriptFullPath = $scriptMapping[$selectedScriptTitle]
    if ($scriptFullPath -and (Test-Path $scriptFullPath)) {
        Remove-Item -Path $scriptFullPath -Force
        RefreshPsScriptsList
    }
})
$contextMenu.Items.Add($deleteMenuItem) | Out-Null

$psScriptsListBox.ContextMenuStrip = $contextMenu
$psScriptsListBox.Add_MouseUp({
    $mouseEventArgs = [System.Windows.Forms.MouseEventArgs]$_
    
    # Determine the item index under the mouse cursor
    $index = $psScriptsListBox.IndexFromPoint($mouseEventArgs.Location)
    
    # If control key is not pressed, clear the current selection
    if (-not [System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Control) {
        $psScriptsListBox.ClearSelected()
    }

    # Set the selected index to the one under the mouse cursor
    if ($index -ge 0) {
        $psScriptsListBox.SelectedIndex = $index
    }
    
    # If it's a right-click, show the context menu
    if ($mouseEventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        if ($index -ge 0) {
            $selectedTitle = $psScriptsListBox.SelectedItem
            # Ensure $selectedTitle is not null before accessing $scriptMapping
            if ($null -ne $selectedTitle) {
                $scriptFullPath = $scriptMapping[$selectedTitle]
                if ($scriptFullPath -and (Test-Path $scriptFullPath)) {
                    $psScriptsListBox.ContextMenuStrip.Show($psScriptsListBox, $mouseEventArgs.Location)
                }
            }
        }
    }
})

$softwareContextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$runSoftwareMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$runSoftwareMenuItem.Text = "Run"
$runSoftwareMenuItem.Add_Click({
    # Code to run the selected software(s) when this option is clicked
    foreach ($selectedSoftware in $softwareListBox.SelectedItems) {
        $softwareFullPath = Join-Path $softwarePath $selectedSoftware
        if ($softwareFullPath -and (Test-Path $softwareFullPath)) {
            Start-Process -FilePath $softwareFullPath -Wait
        }
    }
})

$softwareContextMenu.Items.Add($runSoftwareMenuItem) | Out-Null

$deleteSoftwareMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteSoftwareMenuItem.Text = "Delete"
$deleteSoftwareMenuItem.Add_Click({
    $selectedSoftware = $softwareListBox.SelectedItem
    $softwareFullPath = Join-Path $softwarePath $selectedSoftware
    if ($softwareFullPath -and (Test-Path $softwareFullPath)) {
        Remove-Item -Path $softwareFullPath -Force
        RefreshSoftwareList
    }
})
$softwareContextMenu.Items.Add($deleteSoftwareMenuItem) | Out-Null

$softwareListBox.ContextMenuStrip = $softwareContextMenu

$softwareListBox.Add_MouseUp({
    $mouseEventArgs = [System.Windows.Forms.MouseEventArgs]$_
    
    # Determine the item index under the mouse cursor
    $index = $softwareListBox.IndexFromPoint($mouseEventArgs.Location)
    
    # If control key is not pressed, clear the current selection
    if (-not [System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Control) {
        $softwareListBox.ClearSelected()
    }

    # Set the selected index to the one under the mouse cursor
    if ($index -ge 0) {
        $softwareListBox.SelectedIndex = $index
    }
    
    # If it's a right-click, show the context menu
    if ($mouseEventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        if ($index -ge 0) {
            $softwareListBox.ContextMenuStrip.Show($softwareListBox, $mouseEventArgs.Location)
        }
    }
})
$runAsAdminSoftwareMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$runAsAdminSoftwareMenuItem.Text = "Run as Administrator"
$runAsAdminSoftwareMenuItem.Add_Click({
    # Code to run the software with elevated privileges
    foreach ($selectedSoftware in $softwareListBox.SelectedItems) {
        $softwareFullPath = Join-Path $softwarePath $selectedSoftware
        if ($softwareFullPath -and (Test-Path $softwareFullPath)) {
            Start-Process -FilePath $softwareFullPath -Verb RunAs -Wait
        }
    }
})
$softwareContextMenu.Items.Add($runAsAdminSoftwareMenuItem) | Out-Null

$webShortcutContextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$openShortcutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openShortcutMenuItem.Text = "Open"
$openShortcutMenuItem.Add_Click({
    foreach ($selectedShortcut in $webShortcutsListBox.SelectedItems) {
        $shortcutFullPath = Join-Path $webShortcutsPath "$selectedShortcut.url"
        if ($shortcutFullPath -and (Test-Path $shortcutFullPath)) {
            Start-Process -FilePath $shortcutFullPath
        }
    }
})

$webShortcutContextMenu.Items.Add($openShortcutMenuItem) | Out-Null

$deleteShortcutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteShortcutMenuItem.Text = "Delete"
$deleteShortcutMenuItem.Add_Click({
    # Code to delete the web shortcut when this option is clicked
    $selectedShortcut = $webShortcutsListBox.SelectedItem
    $shortcutFullPath = Join-Path $webShortcutsPath "$selectedShortcut.url"
    if ($shortcutFullPath -and (Test-Path $shortcutFullPath)) {
        Remove-Item -Path $shortcutFullPath -Force
        # Refresh the list after deletion
        RefreshWebShortcutsList
    }
})
$webShortcutContextMenu.Items.Add($deleteShortcutMenuItem) | Out-Null
$runAsAdminWebShortcutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$runAsAdminWebShortcutMenuItem.Text = "Run as Administrator"
$runAsAdminWebShortcutMenuItem.Add_Click({
    # Code to open the web shortcut with elevated privileges
    # Note: This might not be common for URL files, but the functionality is provided for consistency
    $selectedShortcut = $webShortcutsListBox.SelectedItem
    $shortcutFullPath = Join-Path $webShortcutsPath "$selectedShortcut.url"
    if ($shortcutFullPath -and (Test-Path $shortcutFullPath)) {
        Start-Process -FilePath $shortcutFullPath -Verb RunAs
    }
})
$webShortcutContextMenu.Items.Add($runAsAdminWebShortcutMenuItem) | Out-Null

$newShortcutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$newShortcutMenuItem.Text = "New"
$newShortcutMenuItem.Add_Click({
    # Code to prompt user for URL and name
    $urlPrompt = [System.Windows.Forms.MessageBox]::Show("Enter the URL:", "New Web Shortcut", [System.Windows.Forms.MessageBoxButtons]::OKCancel)
    if ($urlPrompt -eq "OK") {
        $namePrompt = [System.Windows.Forms.MessageBox]::Show("Enter the name for the shortcut:", "New Web Shortcut", [System.Windows.Forms.MessageBoxButtons]::OKCancel)
        if ($namePrompt -eq "OK") {
            # Create a new .url file with the provided URL and name
            $shortcutContent = @"
[InternetShortcut]
URL=$urlPrompt
"@
            $shortcutPath = Join-Path $webShortcutsPath "$namePrompt.url"
            $shortcutContent | Out-File $shortcutPath
            # Refresh the WebShortcuts list
            RefreshWebShortcutsList
        }
    }
})
$webShortcutContextMenu.Items.Add($newShortcutMenuItem) | Out-Null


$webShortcutsListBox.ContextMenuStrip = $webShortcutContextMenu

$webShortcutsListBox.Add_MouseUp({
    $mouseEventArgs = [System.Windows.Forms.MouseEventArgs]$_
    
    # Determine the item index under the mouse cursor
    $index = $webShortcutsListBox.IndexFromPoint($mouseEventArgs.Location)
    
    # If control key is not pressed, clear the current selection
    if (-not [System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Control) {
        $webShortcutsListBox.ClearSelected()
    }

    # Set the selected index to the one under the mouse cursor
    if ($index -ge 0) {
        $webShortcutsListBox.SelectedIndex = $index
    }
    
    # If it's a right-click, show the context menu
    if ($mouseEventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        # Ensure the "New" menu item is only available in the WebShortcuts tab
        $newShortcutMenuItem.Visible = ($tabControl.SelectedTab -eq $webShortcutsTab)
        $webShortcutsListBox.ContextMenuStrip.Show($webShortcutsListBox, $mouseEventArgs.Location)
    }
})

# Attach the form's resize event here
$form.Add_Resize({
    AdjustControlPositions
})

# Adjust control positions initially
AdjustControlPositions

# Minimize the PowerShell console window
$consoleWindowHandle = [ConsoleWindow]::GetConsoleWindow()
[ConsoleWindow]::ShowWindow($consoleWindowHandle, 2)  # 2 is for SW_MINIMIZE

# Show the form
$form.ShowDialog() | Out-Null