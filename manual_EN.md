# Dante Preset XML -> Excel Converter Manual

## ■ Overview
A tool for converting Dante Controller preset files (XML) into formatted Excel (.xlsx) files.
Automatically generates device lists, patch matrix, flow information, etc.
Supports AES67 configuration.

## ■ System Requirements
- Windows 10 / 11
- Microsoft Excel installed
- No additional software required

## ■ Files
- `DanteToExcel.ps1`: Main converter tool (PowerShell)
- `DanteToExcel.bat`: Batch file for easy execution

## ■ How to Use
1. Copy `DanteToExcel.ps1` into a folder.
2. Place your Dante preset XML file(s) in the same folder.
3. Right-click `DanteToExcel.ps1` -> "Run with PowerShell" or double-click `DanteToExcel.bat`.
4. Select the output mode (usually `1` for Default).
5. Once complete, an .xlsx file will be created in the same folder.

## ■ Output Modes
- **Default (1)**: Summary mode with essential device and flow information.
- **Detail (2)**: Full mode including all channel-level details and subscriptions.

## ■ Notes
- Excel runs in the background during conversion. Do not manually operate Excel until the process is finished.
- For large channel counts, generating the patch matrix may take some time.
