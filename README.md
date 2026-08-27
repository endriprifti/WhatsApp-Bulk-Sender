# WhatsApp Excel Bulk Sender (PowerShell)

A hardened PowerShell script to automate sending sequential, personalized WhatsApp messages with image/PDF attachments directly from an Excel spreadsheet. 

This script targets the **WhatsApp Desktop Microsoft Store App (WinUI 3)**. It uses UI Automation and advanced Win32 API focus-management to solve the common issues associated with simulating keystrokes in UWP/Sandboxed Windows apps.

## ✨ Features
* **Dynamic Excel Parsing:** Pulls Salutations, Full Names, and File Paths directly from an `.xlsx` file.
* **Smart UI Automation:** Uses `Clipboard::SetFileDropList()` instead of clicking through file explorer dialogs.
* **Win32 Focus Lock Bypass:** Bypasses Windows' foreground-lock restrictions by injecting synthetic ALT key events, guaranteeing that sensitive texts/files are never pasted into the wrong window.
* **Full UTF-8 Emoji Support:** Safely handles complex formatting, line breaks, and emojis (🎬, 📅) by routing text through the clipboard rather than simulating individual keystrokes.
* **Safe Teardown:** Clears the clipboard automatically upon exit or error to ensure privacy.

## 📋 Prerequisites
1. **Windows 10 or 11**
2. **Microsoft Excel** (installed locally, as it uses the Excel COM object).
3. **WhatsApp Desktop** (from the Microsoft Store).
4. An Excel file named `List.xlsx` formatted as follows:
   * **Column A:** Salutation (e.g., *Z.* or *Znj.*)
   * **Column B:** Full Name
   * **Column C:** Absolute Path to the attachment (e.g., `C:\Photographs\photo-1.jpg`)

## 🚀 How to Use

1. **Clone the repository:**
   ```bash
   git clone https://github.com/endriprifti/WhatsApp-Bulk-Sender

## Configure the script:
Open `Send-WhatsAppBulk.ps1` and adjust the `$ExcelFilePath` variable at the top to point to your spreadsheet.

Customize your message:
Edit the `$caption` variable block in the script to match your desired template and language.

Open WhatsApp:
Open the WhatsApp Desktop app and click on the chat/contact you want to dispatch the sequence to. Leave it open on your screen.

Run the script:
Open PowerShell (standard, non-administrator) and run:

## PowerShell

`.\Send-WhatsAppBulk.ps1`

⚠️ Note: Do not touch the mouse or keyboard while the script is actively looping through the Excel rows.

## 🛠 Troubleshooting
`"WhatsApp window not found"`:
Microsoft Store apps hide their main processes. Ensure WhatsApp is not minimized to the system tray and is visible on your desktop before running.

Script skips files / Pastes text as a separate message:
This means your PC is taking longer to load the image into WhatsApp. Increase the `$WaitBeforePaste` variable at the top of the script from 3 to 4 or 5.
