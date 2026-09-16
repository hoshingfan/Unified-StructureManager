# Launch-GUI.ps1 - GUI with config loading AND saving
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configFile = Join-Path $scriptDir "config.json"

# Function to save current GUI settings to config.json
function Save-Config {
    $config = @{
        SourcePath      = $txtSource.Text
        CsvPath         = $txtCsv.Text
        DestinationPath = $txtDest.Text
        CutoffDate      = $dtCutoff.Value.ToString('yyyy-MM-dd')
        EnableCutoff    = $chkUseCutoff.Checked
        Mode            = $cmbMode.SelectedItem.ToString()
        Force           = $chkForce.Checked
        VerboseLogging  = $chkVerbose.Checked
        ExportFiles     = $chkFiles.Checked
        ExportFolders   = $chkFolders.Checked
    }
    $configJson = $config | ConvertTo-Json
    Set-Content -Path $configFile -Value $configJson -Encoding UTF8
    Write-Host "Settings saved to $configFile" -ForegroundColor Green
}

# Load config if exists
$cfg = $null
if (Test-Path $configFile) {
    $cfg = Get-Content $configFile -Raw | ConvertFrom-Json
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Unified Structure Manager"
$form.Size = New-Object System.Drawing.Size(620, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Optionally save when form closes
# $form.Add_FormClosing({ Save-Config })

$y = 20

# Source path
$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Text = "Source Path:"
$labelSource.Location = New-Object System.Drawing.Point(20, $y)
$labelSource.AutoSize = $true
$form.Controls.Add($labelSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(150, $y)
$txtSource.Width = 350
$txtSource.Text = if ($cfg -and $cfg.SourcePath) { $cfg.SourcePath } else { "E:\Deptshare\GroupSupplyChain-nlw" }
$form.Controls.Add($txtSource)

$btnSource = New-Object System.Windows.Forms.Button
$btnSource.Text = "Browse"
$btnSource.Location = New-Object System.Drawing.Point(510, $y)
$btnSource.Width = 70
$btnSource.Add_Click({
    $folder = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folder.ShowDialog() -eq "OK") { $txtSource.Text = $folder.SelectedPath }
})
$form.Controls.Add($btnSource)

$y += 40

# CSV path
$labelCsv = New-Object System.Windows.Forms.Label
$labelCsv.Text = "CSV Path:"
$labelCsv.Location = New-Object System.Drawing.Point(20, $y)
$labelCsv.AutoSize = $true
$form.Controls.Add($labelCsv)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = New-Object System.Drawing.Point(150, $y)
$txtCsv.Width = 350
$txtCsv.Text = if ($cfg -and $cfg.CsvPath) { $cfg.CsvPath } else { "GroupSupplyChain-nlw-before2019.csv" }
$form.Controls.Add($txtCsv)

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = "Browse"
$btnCsv.Location = New-Object System.Drawing.Point(510, $y)
$btnCsv.Width = 70
$btnCsv.Add_Click({
    $save = New-Object System.Windows.Forms.SaveFileDialog
    $save.Filter = "CSV Files|*.csv"
    if ($save.ShowDialog() -eq "OK") { $txtCsv.Text = $save.FileName }
})
$form.Controls.Add($btnCsv)

$y += 40

# Destination path
$labelDest = New-Object System.Windows.Forms.Label
$labelDest.Text = "Destination Path:"
$labelDest.Location = New-Object System.Drawing.Point(20, $y)
$labelDest.AutoSize = $true
$form.Controls.Add($labelDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(150, $y)
$txtDest.Width = 350
$txtDest.Text = if ($cfg -and $cfg.DestinationPath) { $cfg.DestinationPath } else { "E:\Restored" }
$form.Controls.Add($txtDest)

$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Browse"
$btnDest.Location = New-Object System.Drawing.Point(510, $y)
$btnDest.Width = 70
$btnDest.Add_Click({
    $folder = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folder.ShowDialog() -eq "OK") { $txtDest.Text = $folder.SelectedPath }
})
$form.Controls.Add($btnDest)

$y += 40

# Cutoff date
$labelCutoff = New-Object System.Windows.Forms.Label
$labelCutoff.Text = "Cutoff Date (optional):"
$labelCutoff.Location = New-Object System.Drawing.Point(20, $y)
$labelCutoff.AutoSize = $true
$form.Controls.Add($labelCutoff)

$dtCutoff = New-Object System.Windows.Forms.DateTimePicker
$dtCutoff.Location = New-Object System.Drawing.Point(150, $y)
$dtCutoff.Width = 150
$dtCutoff.Format = "Short"
if ($cfg -and $cfg.CutoffDate) {
    try { $dtCutoff.Value = [DateTime]::Parse($cfg.CutoffDate) } catch {}
}
$form.Controls.Add($dtCutoff)

$chkUseCutoff = New-Object System.Windows.Forms.CheckBox
$chkUseCutoff.Text = "Enable filter (LastWriteTime < cutoff)"
$chkUseCutoff.Location = New-Object System.Drawing.Point(320, $y)
$chkUseCutoff.AutoSize = $true
if ($cfg -and $cfg.EnableCutoff) { $chkUseCutoff.Checked = $cfg.EnableCutoff }
$form.Controls.Add($chkUseCutoff)

$y += 40

# Mode
$labelMode = New-Object System.Windows.Forms.Label
$labelMode.Text = "Mode:"
$labelMode.Location = New-Object System.Drawing.Point(20, $y)
$labelMode.AutoSize = $true
$form.Controls.Add($labelMode)

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Location = New-Object System.Drawing.Point(150, $y)
$cmbMode.Width = 120
$cmbMode.DropDownStyle = "DropDownList"
$cmbMode.Items.AddRange(@("Export", "Preview", "Recreate", "Fix"))
if ($cfg -and $cfg.Mode) {
    $idx = $cmbMode.Items.IndexOf($cfg.Mode)
    if ($idx -ge 0) { $cmbMode.SelectedIndex = $idx }
} else { $cmbMode.SelectedIndex = 0 }
$form.Controls.Add($cmbMode)

$y += 40

# Export options (Files and Folders checkboxes)
$groupExport = New-Object System.Windows.Forms.GroupBox
$groupExport.Text = "Export Options (only applies to Export mode)"
$groupExport.Location = New-Object System.Drawing.Point(20, $y)
$groupExport.Size = New-Object System.Drawing.Size(560, 80)
$form.Controls.Add($groupExport)

$chkFiles = New-Object System.Windows.Forms.CheckBox
$chkFiles.Text = "Include Files"
$chkFiles.Location = New-Object System.Drawing.Point(20, 25)
$chkFiles.AutoSize = $true
$chkFiles.Checked = if ($cfg -and $cfg.ExportFiles) { $cfg.ExportFiles } else { $true }
$groupExport.Controls.Add($chkFiles)

$chkFolders = New-Object System.Windows.Forms.CheckBox
$chkFolders.Text = "Include Folders"
$chkFolders.Location = New-Object System.Drawing.Point(200, 25)
$chkFolders.AutoSize = $true
$chkFolders.Checked = if ($cfg -and $cfg.ExportFolders) { $cfg.ExportFolders } else { $true }
$groupExport.Controls.Add($chkFolders)

$y += 90

# Force and Verbose
$chkForce = New-Object System.Windows.Forms.CheckBox
$chkForce.Text = "Force (skip confirmation)"
$chkForce.Location = New-Object System.Drawing.Point(150, $y)
$chkForce.AutoSize = $true
if ($cfg -and $cfg.Force) { $chkForce.Checked = $cfg.Force }
$form.Controls.Add($chkForce)

$chkVerbose = New-Object System.Windows.Forms.CheckBox
$chkVerbose.Text = "Verbose Logging"
$chkVerbose.Location = New-Object System.Drawing.Point(320, $y)
$chkVerbose.AutoSize = $true
if ($cfg -and $cfg.VerboseLogging) { $chkVerbose.Checked = $cfg.VerboseLogging }
$form.Controls.Add($chkVerbose)

$y += 50

# Buttons row
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save Settings"
$btnSave.Location = New-Object System.Drawing.Point(150, $y)
$btnSave.Width = 100
$btnSave.Height = 30
$btnSave.BackColor = "LightBlue"
$btnSave.Add_Click({ Save-Config })
$form.Controls.Add($btnSave)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "RUN"
$btnRun.Location = New-Object System.Drawing.Point(350, $y)
$btnRun.Width = 100
$btnRun.Height = 30
$btnRun.BackColor = "LightGreen"
$btnRun.Add_Click({
    Save-Config   # Auto-save current settings before running
    $params = @()
    $params += "-Mode `"$($cmbMode.SelectedItem)`""
    $params += "-SourcePath `"$($txtSource.Text)`""
    $params += "-CsvPath `"$($txtCsv.Text)`""
    $params += "-DestinationPath `"$($txtDest.Text)`""
    if ($chkUseCutoff.Checked) { $params += "-CutoffDate `"$($dtCutoff.Value.ToString('yyyy-MM-dd'))`"" }
    if ($chkForce.Checked) { $params += "-Force" }
    if ($chkVerbose.Checked) { $params += "-VerboseLogging" }
    if ($chkFiles.Checked) { $params += "-ExportFiles" }
    if ($chkFolders.Checked) { $params += "-ExportFolders" }
    $argString = $params -join " "
    $mainScript = Join-Path $scriptDir "Unified-StructureManager.ps1"
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$mainScript`" $argString"
    $form.Close()
})
$form.Controls.Add($btnRun)

$form.ShowDialog() | Out-Null