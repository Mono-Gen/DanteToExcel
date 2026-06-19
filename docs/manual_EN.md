# Dante Preset XML -> Excel Converter Manual

## ■ Overview
A tool for converting Dante Controller preset files (XML) into formatted Excel (.xlsx) workbooks.
Written in Go, this tool directly generates binary `.xlsx` files without needing Microsoft Excel installed. It runs instantly and cross-platform (Windows & macOS), converting presets, building a routing patch matrix, analyzing AES67 flows, and extracting deep clock/network parameters.

## ■ System Requirements
- **Windows**: Windows 10 / 11 (64-bit)
- **macOS**: macOS 10.15 or later (Intel & Apple Silicon)
- **Spreadsheet Viewer**: Any spreadsheet viewer (Microsoft Excel, LibreOffice, Google Sheets, etc.) to view the output `.xlsx` file.

## ■ Files
- `DanteToExcel_windows_x64.exe` : Windows executable
- `DanteToExcel_macOS_Intel` : macOS (Intel) executable
- `DanteToExcel_macOS_AppleSilicon` : macOS (Apple Silicon) executable

## ■ How to Use
1. Place the appropriate executable file for your OS in a folder.
2. Place your Dante preset XML file(s) in the same folder.
3. Double-click the executable to run it (or execute from a command line).
4. If multiple XML files are found in the directory, you will be prompted to select one by typing its number and pressing `Enter`.
5. Select the output mode (**1** for Default, **2** for Detail).
6. Once complete, an `.xlsx` file will be created in the same folder, and the console will prompt you to press `Enter` to exit.

## ■ Output Modes
- **Default (1)**: Summary mode containing:
  - `Devices` (Essential properties)
  - `Patch Matrix` (Visual routing matrix grid)
  - `TX Flows` (Multicast flow configuration)
- **Detail (2)**: Full mode containing everything from Default, plus:
  - Additional detailed fields in `Devices` (e.g. Pull Up, Gateway, Subnet Masks, PTP v2 priority, domain number)
  - `TX Channels` (Flat list of transmit channels)
  - `RX Channels` (Flat list of receive channels)
  - `Subscriptions` (Connection map table)

## ■ Notes
- If an `.xlsx` file with the same name already exists in the folder, it will be overwritten.
- Since this tool directly outputs spreadsheet files without using COM, it will not interfere with any running Excel instances, and execution is nearly instantaneous (under 1 second).
