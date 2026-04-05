# DanteToExcel

**Dante Preset XML -> Excel Converter Tool**

[![Language](https://img.shields.io/badge/Language-PowerShell-blue.svg)](https://docs.microsoft.com/en-us/powershell/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## Overview
DanteToExcel is a PowerShell-based utility designed to convert **Audinate Dante Controller** preset files (`.xml`) into easy-to-read **Microsoft Excel** (`.xlsx`) workbooks. It's particularly useful for system documentation, channel mapping verification, and AES67 configuration management.

---

## 🇯🇵 日本語ドキュメント
[manual_JP.md](manual_JP.md) をご覧ください。

### 🇺🇸 English Documentation
See [manual_EN.md](manual_EN.md) for detailed instructions.

---

## Key Features
- **Device List Generation**: Detailed properties (IP, Sample Rate, Latency, AES67, PTP status).
- **Patch Matrix View**: A visual routing matrix similar to the Dante Controller interface (up to 512 channels).
- **TX flows & AES67 Analysis**: Identifies multicast flows and AES67 endpoints.
- **Bilingual Results**: Supports both English and Japanese characters in presets.

## Quick Start
1. Download `DanteToExcel.ps1` and place it in a folder with your `.xml` presets.
2. Right-click the `.ps1` file and select **"Run with PowerShell"**.
3. Choose your output mode (**1: Default** or **2: Detail**).
4. Find the generated `.xlsx` file in the same folder.

## System Requirements
- Windows 10 or 11
- Microsoft Excel installed (used for file generation via COM)

---

## License
Provided under the MIT License. Use it at your own risk.
