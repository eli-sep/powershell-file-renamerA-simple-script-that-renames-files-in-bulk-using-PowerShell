Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add Win32 API support to move the FolderBrowserDialog window
if (-not ([type]::GetType("Win32"))) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
"@
}

function Show-GUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "File Renamer"
    $form.Size = New-Object System.Drawing.Size(500, 400)
    $form.StartPosition = "CenterScreen"

    # Folder path label + text + browse button
    $lblFolder = New-Object System.Windows.Forms.Label
    $lblFolder.Text = "Target Folder:"
    $lblFolder.Location = '10,20'
    $lblFolder.Size = '100,20'

    $txtFolder = New-Object System.Windows.Forms.TextBox
    $txtFolder.Location = '110,20'
    $txtFolder.Size = '260,20'

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse"
    $btnBrowse.Location = '380,18'
    $btnBrowse.Size = '80,24'
    $btnBrowse.Add_Click({
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 100
        $timer.Add_Tick({
            $hwnd = [Win32]::FindWindow("#32770", "Browse For Folder")
            if ($hwnd -ne [IntPtr]::Zero) {
                [Win32]::MoveWindow($hwnd, 400, 200, 600, 700, $true)
                $timer.Stop()
            }
        })
        $timer.Start()

        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select the folder containing the files to rename"
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $txtFolder.Text = $folderBrowser.SelectedPath
        }
        $timer.Stop()
    })

    # Dropdown for method
    $lblMethod = New-Object System.Windows.Forms.Label
    $lblMethod.Text = "Renaming Method:"
    $lblMethod.Location = '10,60'
    $lblMethod.Size = '120,20'

    $cmbMethod = New-Object System.Windows.Forms.ComboBox
    $cmbMethod.Location = '140,60'
    $cmbMethod.Size = '320,20'
    $cmbMethod.DropDownStyle = 'DropDownList'
    $cmbMethod.Items.AddRange(@("Rename files by sequence", "Find and replace in file names", "Rename files by property"))

    # Dynamic parameter labels and textboxes
    $lblParam1 = New-Object System.Windows.Forms.Label
    $lblParam1.Location = '10,100'
    $lblParam1.Size = '200,20'

    $txtParam1 = New-Object System.Windows.Forms.TextBox
    $txtParam1.Location = '220,100'
    $txtParam1.Size = '240,20'

    $lblParam2 = New-Object System.Windows.Forms.Label
    $lblParam2.Location = '10,140'
    $lblParam2.Size = '200,20'

    $txtParam2 = New-Object System.Windows.Forms.TextBox
    $txtParam2.Location = '220,140'
    $txtParam2.Size = '240,20'

    $lblParam3 = New-Object System.Windows.Forms.Label
    $lblParam3.Location = '10,180'
    $lblParam3.Size = '200,20'

    $txtParam3 = New-Object System.Windows.Forms.TextBox
    $txtParam3.Location = '220,180'
    $txtParam3.Size = '240,20'

    $cmbMethod.Add_SelectedIndexChanged({
        switch ($cmbMethod.SelectedIndex) {
            0 {
                $lblParam1.Text = "Prefix:"
                $lblParam2.Text = "Suffix:"
                $lblParam3.Text = "Min Digits:"
                $lblParam1.Visible = $lblParam2.Visible = $lblParam3.Visible = $true
                $txtParam1.Visible = $txtParam2.Visible = $txtParam3.Visible = $true
            }
            1 {
                $lblParam1.Text = "Find Text:"
                $lblParam2.Text = "Replace Text:"
                $lblParam1.Visible = $lblParam2.Visible = $true
                $txtParam1.Visible = $txtParam2.Visible = $true
                $lblParam3.Visible = $txtParam3.Visible = $false
            }
            2 {
                $lblParam1.Text = "Property Name:"
                $lblParam1.Visible = $txtParam1.Visible = $true
                $lblParam2.Visible = $txtParam2.Visible = $false
                $lblParam3.Visible = $txtParam3.Visible = $false
            }
        }
    })
    $form.Add_Shown({
        $form.Activate()
        $cmbMethod.SelectedIndex = 0
    })  # Trigger label update after form is shown

    # Run button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Run"
    $btnRun.Location = '140,240'
    $btnRun.Size = '80,30'
    $btnRun.Add_Click({
        $dir = $txtFolder.Text
        if (-not (Test-Path $dir)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid folder.")
            return
        }
        Set-Location -Path $dir
        $files = Get-ChildItem -File

        switch ($cmbMethod.SelectedIndex) {
            0 {  # Rename by sequence
                $prefix = $txtParam1.Text
                $suffix = $txtParam2.Text
                $minDigits = $txtParam3.Text
                $format = "{0:D$minDigits}"
                $counter = 1
                foreach ($file in $files) {
                    $newName = "$prefix" + ($format -f $counter) + "$suffix"
                    Rename-Item $file.FullName -NewName $newName
                    $counter++
                }
            }
            1 {  # Find and replace
                $find = $txtParam1.Text
                $replace = $txtParam2.Text
                foreach ($file in $files) {
                    if ($file.Name -like "*$find*") {
                        $newName = $file.Name -replace $find, $replace
                        Rename-Item $file.FullName -NewName $newName
                    }
                }
            }
            2 {  # Rename by property
                $property = $txtParam1.Text
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace($dir)
                foreach ($file in $files) {
                    $item = $folder.ParseName($file.Name)
                    $match = $false
                    for ($i = 0; $i -lt 300; $i++) {
                        if ($folder.GetDetailsOf($null, $i) -eq $property) {
                            $value = $folder.GetDetailsOf($item, $i)
                            if ($value) {
                                $newName = "$value$($file.Extension)"
                                Rename-Item $file.FullName -NewName $newName
                            }
                            $match = $true
                            break
                        }
                    }
                    if (-not $match) {
                        Write-Host "Property '$property' not found for $($file.Name)"
                    }
                }
            }
        }
        [System.Windows.Forms.MessageBox]::Show("Renaming complete.", "Done")
    })

    # Exit button
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Exit"
    $btnExit.Location = '260,240'
    $btnExit.Size = '80,30'
    $btnExit.Add_Click({ $form.Close() })

    $form.Controls.AddRange(@(
        $lblFolder, $txtFolder, $btnBrowse,
        $lblMethod, $cmbMethod,
        $lblParam1, $txtParam1,
        $lblParam2, $txtParam2,
        $lblParam3, $txtParam3,
        $btnRun, $btnExit
    ))

    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
}

Show-GUI
