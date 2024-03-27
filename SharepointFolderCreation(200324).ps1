################################################
# Date: March 20, 2024
# Author: Oliver Posa
# Description: A PowerShell script to read information off a File/Folder matrix, create the file structure 
#              and add correct permissions.
#
# Instructions:
# 1. Ensure you're running this script with an account that has access to the SharePoint site your're going
#    to be working on 
#
# 2. Enter the details of the site in the Dialogue Box, ensure that the correct matrix is selected. 
#
# 2. Click OK, and wait for the script to complete it's execution
#
# WARNING: This script can take a long time to complete executing. Please ensure that you're able to keep it
#          going for a long time, or you run it on a server.
###############################################

Add-Type -AssemblyName System.Windows.Forms
# Get the installed version of PowerShell
$psVersion = $PSVersionTable.PSVersion

# Check if the major version is at least 7 and the minor version is at least 2
Write-Host "PowerShell Verison Check: "
if ($psVersion.Major -ge 7 -and $psVersion.Minor -ge 2) {
    Write-Host "PASS: PowerShell version is at least 7.2" -ForegroundColor Green
}
else {
    Write-Host "FAIL: PowerShell version must be greater than 7.2 for the PnP PowerShell functions to run! Please upgrade your PowerShell version by following this link: https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows" -ForegroundColor Red
    return;
}

Write-Host "Module Check: "
# Check if PSExcel module is installed
if ((Test-Path -Path ".\deps\PSExcel-master\PSExcel") -and (Get-Command "Import-XLSX" -errorAction SilentlyContinue)) {
    Write-Host "PASS: PSExcel Module has already been imported!" -ForegroundColor Green
}
else {
    Write-Host "INFO: Installing PSExcel Module..." -ForegroundColor Yellow
    
    try {
        $PSExcelURL = "https://codeload.github.com/RamblingCookieMonster/PSExcel/zip/refs/heads/master"
        $filename = "C:/Temp/psexcel.zip" 
        (New-Object System.Net.WebClient).DownloadFile($PSExcelURL, $filename)
        Expand-Archive $filename -DestinationPath "deps" -Force
        Remove-Item $filename
    
        Get-ChildItem "deps\PSExcel-master\PSExcel" -Recurse -File | % {
            Unblock-File -Path $_.FullName
        }
    
        Import-Module ".\deps\PSExcel-master\PSExcel" -Force
        Write-Host "INFO: PSExcel Module has been installed!" -ForegroundColor Yellow
    }
    catch {
        Write-Host "FAIL: PSExcel Module could not be installed! $_" -ForegroundColor Yellow
        return;
    }
}
# Check if PnP.PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    # PnP.PowerShell module is not installed, so install it system-wide
    Write-Host "INFO: PnP.PowerShell module is not installed. Installing..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope AllUsers -Force
    Write-Host "PASS: PnP.PowerShell module installed successfully." -ForegroundColor Green
}
else {
    # PnP.PowerShell module is already installed
    Write-Host "PASS: PnP.PowerShell module is already installed." -ForegroundColor Green
}

if ( (Get-Command "Import-XLSX" -errorAction SilentlyContinue) -and
    (Get-Command "Connect-PnPOnline" -errorAction SilentlyContinue) -and
    (Get-Command "Resolve-PnPFolder" -errorAction SilentlyContinue) -and
    (Get-Command "Get-PnPGroup" -errorAction SilentlyContinue) -and
    (Get-Command "Set-PnPGroupPermissions" -errorAction SilentlyContinue) -and
    (Get-Command "Add-PnPFolder" -errorAction SilentlyContinue) -and
    (Get-Command "Get-PnPFolder" -errorAction SilentlyContinue) -and
    (Get-Command "Set-PnPListItemPermission" -errorAction SilentlyContinue)) {
    Write-Host "All required commands have been detected, script can begin executing."
}
else {
    Write-Host "One or more required commands is not avalible, this script will not be able to execute! Please review the following information, and import the missing dependancies manually!" -ForegroundColor Red

    if (-not (Get-Command "Import-XLSX" -errorAction SilentlyContinue)) {
        Write-Host "Import-XLSX command is not available."
    }
    if (-not (Get-Command "Connect-PnPOnline" -errorAction SilentlyContinue)) {
        Write-Host "Connect-PnPOnline command is not available."
    }
    if (-not (Get-Command "Resolve-PnPFolder" -errorAction SilentlyContinue)) {
        Write-Host "Resolve-PnPFolder command is not available."
    }
    if (-not (Get-Command "Get-PnPGroup" -errorAction SilentlyContinue)) {
        Write-Host "Get-PnPGroup command is not available."
    }
    if (-not (Get-Command "Set-PnPGroupPermissions" -errorAction SilentlyContinue)) {
        Write-Host "Set-PnPGroupPermissions command is not available."
    }
    if (-not (Get-Command "Add-PnPFolder" -errorAction SilentlyContinue)) {
        Write-Host "Add-PnPFolder command is not available."
    }
    if (-not (Get-Command "Get-PnPFolder" -errorAction SilentlyContinue)) {
        Write-Host "Get-PnPFolder command is not available."
    }
    if (-not (Get-Command "Set-PnPListItemPermission" -errorAction SilentlyContinue)) {
        Write-Host "Set-PnPListItemPermission command is not available."
    }

    return;
}

#Set Parameters
$WebURL = "https://monadel.sharepoint.com/sites/DevSite/"
$TopFolderName = "Shared Documents/Test3"
$excelFile = $null;
# Function to handle button click event
function BrowseFile {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select an Excel File"
    $fileDialog.Multiselect = $false
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $result = $fileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $textboxErrorReportLocation.Text = $fileDialog.FileName
    }
}
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "ShareGate Folder Creator"
$form.Size = New-Object System.Drawing.Size(440, 190)
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.ControlBox = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = "CenterScreen"
# Server Address label and text box
$labelServer = New-Object System.Windows.Forms.Label
$labelServer.Text = "Site URL:"
$labelServer.Location = New-Object System.Drawing.Point(10, 20)
$labelServer.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelServer)

$textboxServer = New-Object System.Windows.Forms.TextBox
$textboxServer.Location = New-Object System.Drawing.Point(120, 20)
$textboxServer.Size = New-Object System.Drawing.Size(290, 20)
$textboxServer.Text = "https://monadel.sharepoint.com/sites/changeme"
$form.Controls.Add($textboxServer)
# User List label and text box
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "Site Folder Path:"
$labelUser.Location = New-Object System.Drawing.Point(10, 50)
$labelUser.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelUser)

$textboxUser = New-Object System.Windows.Forms.TextBox
$textboxUser.Location = New-Object System.Drawing.Point(120, 50)
$textboxUser.Size = New-Object System.Drawing.Size(290, 20)
$textboxUser.Text = "Shared Documents"
$form.Controls.Add($textboxUser)
# Error Report Location label and text box
$labelErrorReport = New-Object System.Windows.Forms.Label
$labelErrorReport.Text = "Folder Matrix Path:"
$labelErrorReport.Location = New-Object System.Drawing.Point(10, 80)
$labelErrorReport.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelErrorReport)

$textboxErrorReportLocation = New-Object System.Windows.Forms.TextBox
$textboxErrorReportLocation.Location = New-Object System.Drawing.Point(120, 80)
$textboxErrorReportLocation.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textboxErrorReportLocation)
# Browse button
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(330, 80)
$buttonBrowse.Size = New-Object System.Drawing.Size(75, 23)
$buttonBrowse.Text = "Browse"
$buttonBrowse.Add_Click({ BrowseFile })
$form.Controls.Add($buttonBrowse)
# Cancel button
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(120, 110)
$buttonCancel.Size = New-Object System.Drawing.Size(75, 23)
$buttonCancel.Text = "Cancel"
$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($buttonCancel)
# OK button
$buttonOK = New-Object System.Windows.Forms.Button
$buttonOK.Location = New-Object System.Drawing.Point(330, 110)
$buttonOK.Size = New-Object System.Drawing.Size(75, 23)
$buttonOK.Text = "OK"
$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($buttonOK)
# Show form
$form.BringToFront()
$form.ShowDialog()

# Output values if OK button is clicked
if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $WebURL = $textboxServer.Text
    $TopFolderName = $textboxUser.Text
    $excelFile = $textboxErrorReportLocation.Text
}
else{
    return;
}

# Import Excel File 
try {
    $file = Import-XLSX -Path $excelFile
}
catch {
    return $_;
}

#Connect to PnP Online
$validConnection = $false;
while ($validConnection -eq $false) {
    try {
        # Show-LoginGUI
        Connect-PnPOnline -Url $WebURL -Interactive
        Resolve-PnPFolder -SiteRelativePath $TopFolderName

        Write-Information "Connection successful!"
        $validConnection = $true;
    }
    catch {
        # Define the message box parameters
        $message = "Login failed, would you like to try again?"
        $title = "Login Failed"
        $buttons = [System.Windows.Forms.MessageBoxButtons]::OKCancel
        $icon = [System.Windows.Forms.MessageBoxIcon]::Warning

        # Show the message box
        $result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)

        # Check the result
        if ($result -eq "OK") {
            $validConnection = $false;
        } else {
            return; #Exit
        }
    }
}

# Check groups
$groupCheckData = @()
$file[0].PSObject.Properties | ForEach-Object {
    $groupCheckData += @{
        Key   = $_.Name
        Value = $_.Value
    }
}
$spNames = @()
Get-PnPGroup |  ForEach-Object {
    $spNames += $_.Title
}
for ($i = 5; $i -lt $groupCheckData.Count; $i++) {
    if ($null -ne $groupCheckData[$i].Key -and $groupCheckData[$i].Key -notmatch "<Column") {
        # actual check
        $spGroupName = $groupCheckData[$i].Key

        $illegalChars = "\/[]:|<>+=;,?*''@".ToCharArray()
        foreach ($char in $illegalChars) {
            $spGroupName = $spGroupName -replace [regex]::Escape($char), ""
        }

        if ($spNames -contains $spGroupName) {
            Write-Host "$spGroupName exists" -ForegroundColor Green
        }
        else {
            Write-Host "$spGroupName does not exist and will be created with read access" -ForegroundColor Yellow
            try {
                $newSpGroup = New-PnPGroup -Title $spGroupName
                Set-PnPGroupPermissions -Identity $newSpGroup -AddRole "Read"
            }
            catch {
                if ($_.Exception -match "The specified name is already in use") {
                    Write-Host "$spGroupName threw an error but exists anyway" -ForegroundColor Green
                }
                else {
                    Write-Host "$($spGroupName): $($_.Exception)" -ForegroundColor Red
                }
            }

        }
    }
}

$folderLocation = $null;
$lvl2Parent = $null;
$lvl3Parent = $null;
$lvl4Parent = $null;
$ontoFileRows = $false;

$fileData
foreach ($row in $file) {
    # Convert custom object to array of key-value pairs
    $rowDataArray = @()
    $row.PSObject.Properties | ForEach-Object {
        $rowDataArray += @{
            Key   = $_.Name
            Value = $_.Value
        }
    }
 
    # Line valididty check
    if ($rowDataArray[0].Value -contains "Lvl 1" -and
        $rowDataArray[1].Value -contains "Lvl 2" -and
        $rowDataArray[2].Value -contains "Lvl 3" -and
        $rowDataArray[3].Value -contains "Lvl 4") {
        $ontoFileRows = $true;
        continue;
    }
 
    if ($ontoFileRows) {
        try {
            # Check column for value
            if ($null -ne $rowDataArray[0].Value) {
                # This is a LVL 1 Folder 
                $folderName = "$($rowDataArray[0].Value) $($rowDataArray[4].Value)"
                $folderLocation = "$TopFolderName/$folderName"
                $lvl2Parent = $folderName # This may be changed to folderName, see how it goes
                Add-PnPFolder -Name $folderName -Folder $TopFolderName | Out-Null
                Write-Host "Folder created at $folderLocation" -ForegroundColor Green
            }
            elseif ($null -ne $rowDataArray[1].Value) {
                # This is a LVL 2 Folder
                $folderName = "$($rowDataArray[1].Value) $($rowDataArray[4].Value)"
                $folderLocation = "$TopFolderName/$lvl2Parent/$folderName"
                $lvl3Parent = $folderName
                Add-PnPFolder -Name $folderName -Folder "$TopFolderName/$lvl2Parent" | Out-Null
                Write-Host "Folder created at $folderLocation" -ForegroundColor Green
            }
            elseif ($null -ne $rowDataArray[2].Value) {
                # This is a LVL 3 Folder  
                $folderName = "$($rowDataArray[2].Value) $($rowDataArray[4].Value)"
                $folderLocation = "$TopFolderName/$lvl2Parent/$lvl3Parent/$folderName"
                $lvl4Parent = $folderName
                Add-PnPFolder -Name $folderName -Folder "$TopFolderName/$lvl2Parent/$lvl3Parent" | Out-Null
                Write-Host "Folder created at $folderLocation" -ForegroundColor Green
            }
            elseif ($null -ne $rowDataArray[3].Value) {
                # This is a LVL 4 Folder
                $folderName = "$($rowDataArray[3].Value) $($rowDataArray[4].Value)"
                # No parent to be set.
                $folderLocation = "$TopFolderName/$lvl2Parent/$lvl3Parent/$lvl4Parent/$folderName"
                Add-PnPFolder -Name $folderName -Folder "$TopFolderName/$lvl2Parent/$lvl3Parent/$lvl4Parent" | Out-Null
                Write-Host "Folder created at $folderLocation" -ForegroundColor Green
            }

            # Break inherited permissions
            Set-PnPList -Identity $folderLocation -BreakRoleInheritance | Out-Null
            # Get all role assignments for the folder
            $Folder = Get-PnPFolder -Url $folderLocation
            $Folder.ListItemAllFields.BreakRoleInheritance($false, $false)
            $RoleAssignments = Get-PnPProperty -ClientObject $Folder.ListItemAllFields -Property RoleAssignments
            # Create a list to store role assignments that need to be deleted
            $RoleAssignmentsToDelete = @()
            # Loop through each role assignment and add it to the list if it's a group
            foreach ($RoleAssignment in $RoleAssignments) {
                $Member = Get-PnPProperty -ClientObject $RoleAssignment -Property Member
                if ($Member.PrincipalType -eq "SharePointGroup") {
                    $RoleAssignmentsToDelete += $RoleAssignment
                }
            }
            # Loop through the list of role assignments to delete and remove them
            foreach ($RoleAssignment in $RoleAssignmentsToDelete) {
                $RoleAssignment.DeleteObject()
                Invoke-PnPQuery
            }
 
            # Set new Permissions
            for ($i = 5; $i -lt $rowDataArray.Count; $i++) {
                if ($null -ne $rowDataArray[$i].Value -and ($rowDataArray[$i].Value -eq "W" -or $rowDataArray[$i].Value -eq "R")) {
                    # Check if cell contains a W or a R, any other data is not to be considered. 
                    $Folder = Get-PnPFolder -Url $folderLocation

                    $spGroupName = $rowDataArray[$i].Key
                    $illegalChars = "\/[]:|<>+=;,?*''@".ToCharArray()
                    foreach ($char in $illegalChars) {
                        $spGroupName = $spGroupName -replace [regex]::Escape($char), ""
                    }
                    try {
                        switch ($rowDataArray[$i].Value) {
                            "W" {     
                                Set-PnPListItemPermission -List "Documents" -Identity $Folder.ListItemAllFields -Group $spGroupName -AddRole 'Contribute'
                                Write-Host "WRITE access provided to $spGroupName" -ForegroundColor DarkGreen
                            }
                            "R" {  
                                Set-PnPListItemPermission -List "Documents" -Identity $Folder.ListItemAllFields -Group $spGroupName -AddRole 'Read'
                                Write-Host "READ access provided to $spGroupName" -ForegroundColor DarkGreen
                            }
                            Default { }
                        }
                    }
                    catch {
                        Write-Host "Failed writing permissions for $($spGroupName): $_" -ForegroundColor Red
                    }
                }
            }
        }
        catch {
            Write-Host "Failed creating folder at $($folderLocation): $_" -ForegroundColor Red
        }
    }
}