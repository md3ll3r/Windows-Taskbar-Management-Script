<#
.SYNOPSIS
    Windows 11 Taskbar Manager v3 (Complete Suite)
.DESCRIPTION
    Features:
    1. Check items to pin to Taskbar.
    2. "Browse" to add custom .exe files.
    3. "Remove" to delete custom apps from the list.
    4. "Up/Down" to reorder icons.
    5. Portable: Saves config to 'saved_apps.json' in the script folder.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# --- 1. Setup Environment & Paths ---
try {
    if ($PSScriptRoot) { $BaseDir = $PSScriptRoot } 
    else { $BaseDir = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') }
} catch { $BaseDir = "$env:APPDATA\TaskbarManagerShortcuts" }

$ConfigPath = "$BaseDir\saved_apps.json"
$ShortcutStorage = "$env:APPDATA\TaskbarManagerShortcuts\Links"

if (-not (Test-Path $ShortcutStorage)) { New-Item -ItemType Directory -Force -Path $ShortcutStorage | Out-Null }

$Global:CustomApps = @()
$DefaultApps = @(
    @{ Name = "Google Chrome";    Path = "C:\Program Files\Google\Chrome\Application\chrome.exe" },
    @{ Name = "Microsoft Edge";   Path = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" },
    @{ Name = "Microsoft Word";   Path = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" },
    @{ Name = "Microsoft Excel";  Path = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" },
    @{ Name = "Microsoft PowerPoint"; Path = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" },
    @{ Name = "Microsoft Outlook";Path = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" },
    @{ Name = "Teams (Classic)";  Path = "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe" },
    @{ Name = "File Explorer";    Path = "C:\Windows\explorer.exe" },
    @{ Name = "Command Prompt";   Path = "C:\Windows\System32\cmd.exe" },
    @{ Name = "Notepad";          Path = "C:\Windows\System32\notepad.exe" }
)

# --- 2. Data Management Functions ---

function Load-Config {
    if (Test-Path $ConfigPath) {
        try {
            $json = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
            # Clear current custom list to avoid duplicates on reload
            $Global:CustomApps = @()
            foreach ($item in $json) { $Global:CustomApps += @{ Name = $item.Name; Path = $item.Path } }
        } catch {}
    }
}

function Save-Config {
    try { $Global:CustomApps | ConvertTo-Json | Set-Content -Path $ConfigPath -Force } catch {}
}

function Create-Shortcut ($TargetExe, $Name) {
    $TargetExe = [System.Environment]::ExpandEnvironmentVariables($TargetExe)
    if (-not (Test-Path $TargetExe)) { return $null }
    
    $LinkPath = "$ShortcutStorage\$Name.lnk"
    try {
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($LinkPath)
        $Shortcut.TargetPath = $TargetExe
        $Shortcut.Save()
        return $LinkPath
    } catch { return $null }
}

function Apply-Layout ($OrderedLinks) {
    $middle = ""
    foreach ($link in $OrderedLinks) {
        $middle += "        <taskbar:DesktopApp DesktopApplicationLinkPath=`"$link`" />`n"
    }
    
    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
$middle
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@

    $LayoutPath = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml"
    Set-Content -Path $LayoutPath -Value $xml -Encoding UTF8
    
    $RegKey1 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
    $RegKey2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    if (Test-Path $RegKey1) { Get-ChildItem $RegKey1 | Where-Object { $_.Name -like "*taskbar*" } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $RegKey2) { Remove-Item $RegKey2 -Recurse -Force -ErrorAction SilentlyContinue }

    Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Start-Process "explorer.exe"
}

Load-Config

# --- 3. GUI Construction ---

$form = New-Object System.Windows.Forms.Form
$form.Text = "Taskbar Manager v3"
$form.Size = New-Object System.Drawing.Size(510, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(15, 15)
$label.Size = New-Object System.Drawing.Size(450, 20)
$label.Text = "Manage Taskbar Icons (Drag/Move to reorder):"
$label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($label)

# List Box
$checkedList = New-Object System.Windows.Forms.CheckedListBox
$checkedList.Location = New-Object System.Drawing.Point(15, 45)
$checkedList.Size = New-Object System.Drawing.Size(360, 380)
$checkedList.CheckOnClick = $true
$checkedList.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Helper to Fill List
function Refresh-List {
    $checkedList.Items.Clear()
    foreach ($app in $DefaultApps) { $checkedList.Items.Add($app.Name) }
    foreach ($app in $Global:CustomApps) {
        $exists = Test-Path ([System.Environment]::ExpandEnvironmentVariables($app.Path))
        if ($exists) { $checkedList.Items.Add($app.Name) } 
        else { $checkedList.Items.Add("$($app.Name) [MISSING]") }
    }
}
Refresh-List
$form.Controls.Add($checkedList)

# --- SIDE BUTTONS (Ordering & Removal) ---

$btnUp = New-Object System.Windows.Forms.Button
$btnUp.Location = New-Object System.Drawing.Point(385, 45)
$btnUp.Size = New-Object System.Drawing.Size(95, 40)
$btnUp.Text = "Move Up"
$btnUp.Add_Click({
    if ($checkedList.SelectedIndex -gt 0) {
        $idx = $checkedList.SelectedIndex
        $item = $checkedList.Items[$idx]
        $isChecked = $checkedList.GetItemChecked($idx)
        
        # Swap Visual
        $checkedList.Items[$idx] = $checkedList.Items[$idx - 1]
        $checkedList.SetItemChecked($idx, $checkedList.GetItemChecked($idx - 1))
        $checkedList.Items[$idx - 1] = $item
        $checkedList.SetItemChecked($idx - 1, $isChecked)
        $checkedList.SelectedIndex = $idx - 1
    }
})
$form.Controls.Add($btnUp)

$btnDown = New-Object System.Windows.Forms.Button
$btnDown.Location = New-Object System.Drawing.Point(385, 95)
$btnDown.Size = New-Object System.Drawing.Size(95, 40)
$btnDown.Text = "Move Down"
$btnDown.Add_Click({
    if ($checkedList.SelectedIndex -ge 0 -and $checkedList.SelectedIndex -lt $checkedList.Items.Count - 1) {
        $idx = $checkedList.SelectedIndex
        $item = $checkedList.Items[$idx]
        $isChecked = $checkedList.GetItemChecked($idx)
        
        # Swap Visual
        $checkedList.Items[$idx] = $checkedList.Items[$idx + 1]
        $checkedList.SetItemChecked($idx, $checkedList.GetItemChecked($idx + 1))
        $checkedList.Items[$idx + 1] = $item
        $checkedList.SetItemChecked($idx + 1, $isChecked)
        $checkedList.SelectedIndex = $idx + 1
    }
})
$form.Controls.Add($btnDown)

# -- REMOVE BUTTON --
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Location = New-Object System.Drawing.Point(385, 160)
$btnRemove.Size = New-Object System.Drawing.Size(95, 40)
$btnRemove.Text = "Remove"
$btnRemove.ForeColor = "Red"
$btnRemove.Add_Click({
    $idx = $checkedList.SelectedIndex
    if ($idx -lt 0) { return }

    $name = $checkedList.Items[$idx] -replace " \[MISSING\]", ""
    
    # Check if it is a Custom App
    $customItem = $Global:CustomApps | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    
    if ($customItem) {
        # Confirm Deletion
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove '$name' from the list?", "Remove App", "YesNo", "Warning")
        if ($confirm -eq "Yes") {
            # Remove from Memory
            $Global:CustomApps = $Global:CustomApps | Where-Object { $_.Name -ne $name }
            # Update File
            Save-Config
            # Update GUI
            $checkedList.Items.RemoveAt($idx)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("You cannot remove built-in Default apps.`n(Only custom added apps can be removed)", "Restricted", "OK", "Information")
    }
})
$form.Controls.Add($btnRemove)

# --- BOTTOM BUTTONS ---

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(15, 435)
$btnBrowse.Size = New-Object System.Drawing.Size(120, 35)
$btnBrowse.Text = "Browse..."
$btnBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Executables (*.exe)|*.exe"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $filePath = $openFileDialog.FileName
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
        $friendlyName = [Microsoft.VisualBasic.Interaction]::InputBox("Shortcut Name:", "Name", $fileName)
        
        if (-not [string]::IsNullOrWhiteSpace($friendlyName)) {
            $Global:CustomApps += @{ Name = $friendlyName; Path = $filePath }
            Save-Config
            $idx = $checkedList.Items.Add($friendlyName)
            $checkedList.SetItemChecked($idx, $true)
            # Scroll to new item
            $checkedList.SelectedIndex = $idx
        }
    }
})
$form.Controls.Add($btnBrowse)

$btnUncheck = New-Object System.Windows.Forms.Button
$btnUncheck.Location = New-Object System.Drawing.Point(145, 435)
$btnUncheck.Size = New-Object System.Drawing.Size(120, 35)
$btnUncheck.Text = "Uncheck All"
$btnUncheck.Add_Click({
    for ($i=0; $i -lt $checkedList.Items.Count; $i++) { $checkedList.SetItemChecked($i, $false) }
})
$form.Controls.Add($btnUncheck)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Location = New-Object System.Drawing.Point(15, 490)
$btnApply.Size = New-Object System.Drawing.Size(465, 50)
$btnApply.Text = "APPLY LAYOUT (Restarts Explorer)"
$btnApply.BackColor = "LightGreen"
$btnApply.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnApply.Add_Click({
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        [System.Windows.Forms.MessageBox]::Show("Please run as Administrator.", "Error", "OK", "Error")
        return
    }

    $OrderedQueue = @()
    $FullList = $DefaultApps + $Global:CustomApps

    # Read items in their current Visual Order
    for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
        if ($checkedList.GetItemChecked($i)) {
            $itemName = $checkedList.Items[$i] -replace " \[MISSING\]", ""
            
            $appData = $FullList | Where-Object { $_.Name -eq $itemName } | Select-Object -First 1
            if ($appData) {
                $lnk = Create-Shortcut -TargetExe $appData.Path -Name $appData.Name
                if ($lnk) { $OrderedQueue += $lnk }
            }
        }
    }

    $result = [System.Windows.Forms.MessageBox]::Show("Explorer will now restart to apply changes.`nContinue?", "Confirm", "YesNo", "Question")
    if ($result -eq "Yes") { Apply-Layout -OrderedLinks $OrderedQueue }
})
$form.Controls.Add($btnApply)

$form.ShowDialog()