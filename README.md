# DanteToExcel

**Dante Preset XML -> Excel Converter Tool**

[![Language](https://img.shields.io/badge/Language-Go-blue.svg)](https://go.dev/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## Overview
DanteToExcel is a Go-based utility designed to convert **Audinate Dante Controller** preset files (`.xml`) into easy-to-read **Microsoft Excel** (`.xlsx`) workbooks. Rewritten in Go, it features lightning-fast execution, zero external dependencies (no Microsoft Excel installation required), and cross-platform support.

---

## 🇯🇵 日本語ドキュメント (Japanese Manual)
[docs/manual_JP.md](docs/manual_JP.md) をご覧ください。

### 🇺🇸 English Documentation
See [docs/manual_EN.md](docs/manual_EN.md) for detailed instructions.

---

## Key Features
- **Zero Excel Dependency**: Directly generates binary `.xlsx` files. Microsoft Excel does not need to be installed on the system.
- **Vastly Improved Speed**: Converts presets in under 1 second, compared to minutes in the previous PowerShell COM-based script.
- **Detailed Network & Clock Parsing**: Supports PTP v2 domain numbers, DSCP, TTL values, pull-up values, and interface subnet masks and gateways in Detail mode.
- **Visual Patch Matrix View**: Generates a grid showing connection points highlighted in light green, with frozen headers for easy scrolling (supports up to 512 channels).
- **TX Flows & AES67 Analysis**: Automatically detects AES67 multicast flows and details session/transport parameters.
- **Cross-Platform**: Supports Windows and macOS (Intel & Apple Silicon).

## Project Directory Structure
- `src/`: Core Go source code (`main.go`, `main_test.go`).
- `docs/`: User manual documents (`manual_JP.md`, `manual_EN.md`).
- `scripts/`: Build utility script (`build_release.bat`).
- `dist/`: Output directory containing compiled binaries and packaged release Zip archives.

## Quick Start (For Users)
1. Download the release Zip package matching your OS from the releases or `dist/` directory.
2. Extract the archive and place the executable (`DanteToExcel_windows_x64.exe` or `DanteToExcel_macOS_*`) in a folder containing your `.xml` preset files.
3. Simply double-click the executable to launch (or run from a terminal).
4. Select the preset file and preferred output mode (**1: Default** or **2: Detail**).
5. The generated `.xlsx` workbook will be created in the same folder.

## Build from Source (For Developers)
If you have Go installed on your system, you can build the binaries yourself:

```bash
# Run unit tests
go test -v ./src/...

# Package and cross-compile for all supported platforms (Windows, macOS)
# Run from cmd or PowerShell
.\scripts\build_release.bat
```

## System Requirements
- **Windows**: Windows 10 / 11 (64-bit)
- **macOS**: macOS 10.15 or later (Intel & Apple Silicon)
- **Excel Viewer**: Any spreadsheet editor (Microsoft Excel, LibreOffice, Google Sheets, etc.) to view the output `.xlsx` file.

---

## License
Provided under the MIT License. Use it at your own risk.
