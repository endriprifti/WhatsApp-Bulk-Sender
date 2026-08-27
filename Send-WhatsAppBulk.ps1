<#
.SYNOPSIS
    Automates sending personalized WhatsApp messages with file attachments using data from Excel.

.DESCRIPTION
    This script reads a list of recipients and file paths from an Excel spreadsheet.
    It targets the Windows Store version of WhatsApp Desktop (WinUI 3) using UI Automation
    and Win32 API focus management to reliably simulate Drag & Drop / Copy & Paste operations.
    
    It bypasses Windows foreground-lock restrictions by injecting a synthetic ALT key event 
    before enforcing window focus, ensuring the script does not paste sensitive data into the wrong window.

.PREREQUISITES
    - Windows 10 or Windows 11
    - WhatsApp Desktop (Microsoft Store version)
    - Microsoft Excel installed locally
    - An Excel file with 3 columns: Salutation (A), Full Name (B), File Path (C)

.NOTES
    Author: Endri Prifti
    Date: August 2026
    License: MIT
    Warning: Use responsibly and comply with WhatsApp's Terms of Service regarding automated messaging.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient

# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================
$ExcelFilePath = "C:\List.xlsx" # Default path (change as needed)

# Delays to accommodate slower PCs or large attachments (in seconds)
$WaitBeforePaste = 3
$WaitBeforeNext  = 3

# ==============================================================================
# 2. WIN32 API DEFINITIONS (Focus Management)
# ==============================================================================
if (-not ([System.Management.Automation.PSTypeName]'Win32Focus').Type) {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32Focus {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    }
"@
}

# ==============================================================================
# 3. HELPER FUNCTIONS
# ==============================================================================
function Ensure-WhatsAppFocus([IntPtr]$hWnd) {
    # Send dummy ALT keydown/keyup to unlock Windows foreground privilege
    [Win32Focus]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
    [Win32Focus]::keybd_event(0x12, 0, 0x0002, [UIntPtr]::Zero)
    
    [Win32Focus]::SetForegroundWindow($hWnd) | Out-Null
    Start-Sleep -Milliseconds 300
    
    return ([Win32Focus]::GetForegroundWindow() -eq $hWnd)
}

# ==============================================================================
# 4. EXCEL DATA EXTRACTION
# ==============================================================================
$itemsToSend = @()
$excel    = $null
$workbook = $null
$sheet    = $null

try {
    Write-Host "Reading data from: $ExcelFilePath..." -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($ExcelFilePath)
    $sheet    = $workbook.Sheets.Item(1)

    $lastRow = $sheet.Cells($sheet.Rows.Count, "B").End(-4162).Row

    for ($row = 2; $row -le $lastRow; $row++) {
        $title    = ([string]$sheet.Cells.Item($row, 1).Value2).Trim()
        $fullName = ([string]$sheet.Cells.Item($row, 2).Value2).Trim()
        $filePath = ([string]$sheet.Cells.Item($row, 3).Value2).Trim()

        if ([string]::IsNullOrWhiteSpace($fullName) -or [string]::IsNullOrWhiteSpace($filePath)) {
            continue
        }

        # --- DYNAMIC PAYLOAD GENERATION ---
        # Modify this section to change the message logic and language
        if ($title -eq "Mr." -or $title -eq "Mr") {
            $salutation = "Honorate Mr. $fullName"
        } elseif ($title -eq "Mrs." -or $title -eq "Mrs") {
            $salutation = "Honorate Mrs. $fullName"
        } else {
            $salutation = "Honorate Mr./Mrs. $fullName"
        }

        $caption = @"
$salutation,

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
"@

        $itemsToSend += [PSCustomObject]@{
            Title    = $title
            FullName = $fullName
            FilePath = $filePath
            Caption  = $caption
        }
    }
}
finally {
    # Ensure COM objects are fully released to prevent background Excel processes
    if ($workbook) { $workbook.Close($false) }
    if ($excel)    { $excel.Quit() }
    if ($sheet)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null }
    if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($itemsToSend.Count -eq 0) {
    Write-Warning "No valid rows found in the Excel file."
    return
}

# ==============================================================================
# 5. WHATSAPP UI AUTOMATION
# ==============================================================================
# Locate WhatsApp window via UI Automation (Bypasses UWP Process Handle limitations)
$desktop = [System.Windows.Automation.AutomationElement]::RootElement
$allWindows = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)

$waElement = $null
foreach ($window in $allWindows) {
    if ($window.Current.Name -match "WhatsApp") {
        $waElement = $window
        break
    }
}

if (-not $waElement) {
    Write-Error "WhatsApp window not found. Please ensure the app is open on your screen."
    return
}

$waHwnd = [IntPtr]$waElement.Current.NativeWindowHandle
$sentCount    = 0
$skippedCount = 0

try {
    Write-Host "Starting automation for $($itemsToSend.Count) items..." -ForegroundColor Cyan

    foreach ($item in $itemsToSend) {
        $filePath = $item.FilePath
        $caption  = $item.Caption
        $fullName = $item.FullName
        $fileName = Split-Path -Path $filePath -Leaf

        if (-not (Test-Path -Path $filePath -PathType Leaf)) {
            Write-Warning "File not found: '$filePath' for $fullName. Skipping."
            $skippedCount++
            continue
        }

        # Step 1: Stage the file into the Windows Clipboard as a FileDropList
        $fileList = New-Object System.Collections.Specialized.StringCollection
        $fileList.Add((Resolve-Path $filePath).Path)
        [System.Windows.Forms.Clipboard]::SetFileDropList($fileList)

        # Step 2: Ensure WhatsApp is focused before pasting the file
        if (-not (Ensure-WhatsAppFocus -hWnd $waHwnd)) {
            Write-Warning "Could not gain focus for: $fullName. Skipping."
            $skippedCount++
            continue
        }

        # Step 3: Trigger file paste (opens WhatsApp's media preview modal)
        [System.Windows.Forms.SendKeys]::SendWait("^v")

        # Wait for the UI to render the preview and focus the caption box
        Start-Sleep -Seconds $WaitBeforePaste

        # Step 4: Validate focus, paste text, and dispatch
        if (Ensure-WhatsAppFocus -hWnd $waHwnd) {
            [System.Windows.Forms.Clipboard]::SetText($caption)
            [System.Windows.Forms.SendKeys]::SendWait("^v")
            Start-Sleep -Milliseconds 500
            
            [System.Windows.Forms.SendKeys]::SendWait("~") # Enter key
            Write-Host "Successfully sent to: $($item.Title) $fullName ($fileName)" -ForegroundColor Green
            $sentCount++
        } else {
            Write-Error "Focus lost during text paste for $fullName! Aborting this message."
            [System.Windows.Forms.SendKeys]::SendWait("{ESC}") # Attempt to close the dangling preview modal
            $skippedCount++
        }

        # Wait for the modal to unmount before looping to the next file
        Start-Sleep -Seconds $WaitBeforeNext
    }
}
finally {
    # Step 5: Clear clipboard to prevent sensitive data/files from lingering
    [System.Windows.Forms.Clipboard]::Clear()
    Write-Host "`nExecution completed! Sent: $sentCount | Skipped/Errors: $skippedCount" -ForegroundColor Cyan
}