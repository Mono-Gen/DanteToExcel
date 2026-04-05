# === ヘルプ表示 ===
$HelpText = @"
================================================================================
  Dante Preset XML -> Excel Converter
  User Guide
================================================================================

  OVERVIEW
  --------
  Converts Dante Controller preset files (XML) into Excel (.xlsx).
  Generates device list, patch matrix, flow information, etc.
  Supports AES67 configuration.


  REQUIREMENTS
  ------------
  - Windows 10 / 11
  - Microsoft Excel installed
  - No additional software required


  FILE
  ----
  Only one file is needed:

    DanteToExcel.ps1    ... Converter tool


  HOW TO USE
  ----------
  1. Place DanteToExcel.ps1 in any folder.

  2. Place the Dante preset XML file(s) in the same folder.

       Example:
         C:\Dante\
           DanteToExcel.ps1
           MyPreset.xml

  3. Right-click DanteToExcel.ps1 -> "Run with PowerShell"
     Or double-click to run.

  4. A menu appears:

       === Menu ===
         1: Default (summary)
         2: Detail  (all info)
         H: Help

       Select (1 / 2 / H) [default: 1]:

     - Enter 1 or 2 to select the output mode.
     - Press Enter only to use Default mode.
     - Enter H to display this help guide.

  5. If multiple XML files exist, a selection menu appears.

  6. An .xlsx file is created in the same folder.


  OUTPUT MODE COMPARISON
  ----------------------

    Sheet            | Default  | Detail
    -----------------+----------+---------
    Devices          | O 19col  | O 36col
    Patch Matrix     | O        | O
    TX Flows         | O 6col   | O 11col
    TX Channels      | -        | O
    RX Channels      | -        | O
    Subscriptions    | -        | O


  SHEET DETAILS
  -------------

  [Devices] - Device List
    Preset info (name, description, version) shown at top.
    All cells are text format (prevents long ID display issues).

    Default (19 columns):
      Device Name, Default Name, Friendly Name, Model, Manufacturer,
      Model Version, Device Type String, Sample Rate, Encoding,
      Latency (us), Redundancy, External Word Clock,
      Pri IPv4 Address, Pri IPv4 Mode, Sec IPv4 Address, Sec IPv4 Mode,
      Preferred Master, Interop Mode, Clock Preferred

    Detail adds (total 36 columns):
      Manufacturer ID, Model ID, Device Type, Device ID, Process ID,
      Pri Network, Sec Network, Switch VLAN, AES67 MC Prefix,
      Clock Subdomain, PTP v1/v2 Enabled, PTP v1/v2 Unicast Delay,
      Clock Follower Only, TX Ch, RX Ch


  [Patch Matrix] - Routing Matrix
    Visual routing matrix similar to Dante Controller (up to 512ch).
    - Columns = TX channels (device name + label, vertical text)
    - Rows    = RX channels (device name + channel name)
    - Green cell = connected (subscription exists)
    - Headers are frozen (freeze panes enabled)
    - Row height auto-adjusts to fit device name length


  [TX Flows] - Transmit Flow List
    Flow Type column identifies the flow type:
    - AES67 : has destination address + FPP + transportType=2
    - Dante : all others

    Default (6 columns):
      Device Name, Flow Type, Dest Address, Dest Port,
      Slot Count, Slot Channels

    Detail adds (total 11 columns):
      Dante ID, FPP, Media Type, Session ID, Transport Type


  [TX Channels] - Transmit Channel List (Detail only)
    Device Name, Dante ID, Channel Label, Media Type


  [RX Channels] - Receive Channel List (Detail only)
    Device Name, Dante ID, Channel Name, Media Type,
    Subscribed Device, Subscribed Channel, Status


  [Subscriptions] - Connection List (Detail only)
    No., RX Device, RX Channel, RX Dante ID,
    TX Device, TX Channel, Media Type


  SUPPORTED XML TAGS
  ------------------
    Device:  name, default_name, friendly_name, model_name,
      manufacturer_name/id, model_id, model_version,
      device_type/string, instance_id, samplerate, encoding,
      unicast_latency, redundancy, external_word_clock,
      switch_vlan, preferred_master

    Network: interface (network, ipv4_address)
      * Supports primary + secondary interfaces

    AES67:   rtp (interop_mode, aes67_multicast_address_prefix)

    Clock:   clock (subdomain_name, v1/v2_enabled,
      v1/v2_unicast_delay_requests)
      clock_priority (preferred, follower_only)

    Channel: txchannel, rxchannel, txflow


  NOTES
  -----
  - Existing .xlsx with the same name will be overwritten.
  - Excel runs in background during conversion.
    Do not manually operate Excel until completion.
  - If an error occurs, Excel process may remain.
    Use Task Manager to end "Microsoft Excel" if needed.
  - Large channel counts (up to 512ch) may take longer.


  VERSION
  -------
  Supported XML: Dante Controller preset (version 2.1.0 / 3.0.0)

================================================================================
"@

# === 自己起動 ===
if ($host.Name -eq 'ConsoleHost' -and -not $env:DANTE_LAUNCHED) {
    $env:DANTE_LAUNCHED = "1"
    $psArgs = '-ExecutionPolicy Bypass -NoProfile -File "{0}"' -f $MyInvocation.MyCommand.Definition
    Start-Process powershell.exe -ArgumentList $psArgs -NoNewWindow -Wait
    exit
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}
if ([string]::IsNullOrEmpty($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Dante Preset XML -> Excel Converter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Script folder: $ScriptDir" -ForegroundColor Gray
Write-Host ""

# === モード選択 ===
Write-Host "=== Menu ===" -ForegroundColor Cyan
Write-Host "  1: Default (summary)" -ForegroundColor White
Write-Host "  2: Detail  (all info)" -ForegroundColor White
Write-Host "  H: Help" -ForegroundColor White
Write-Host ""
$modeChoice = Read-Host "Select (1 / 2 / H) [default: 1]"

if ($modeChoice -eq "H" -or $modeChoice -eq "h") {
    Write-Host ""
    Write-Host $HelpText
    Write-Host ""
    $modeChoice = Read-Host "Continue? Select mode (1 or 2) [default: 1]"
}

if ($modeChoice -eq "2") {
    $detailMode = $true
    Write-Host "-> Detail mode" -ForegroundColor Yellow
} else {
    $detailMode = $false
    Write-Host "-> Default mode" -ForegroundColor Green
}
Write-Host ""

# === XML ファイル検出 ===
$xmlFiles = @(Get-ChildItem -Path "$ScriptDir\*.xml" -File -ErrorAction SilentlyContinue)

if ($xmlFiles.Count -eq 0) {
    Write-Host "[ERROR] .xml file not found in:" -ForegroundColor Red
    Write-Host "  $ScriptDir" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Contents of folder:" -ForegroundColor Gray
    Get-ChildItem -Path $ScriptDir | ForEach-Object { Write-Host "  $($_.Name)" }
    Read-Host "Enter to exit"; exit 1
}

if ($xmlFiles.Count -eq 1) {
    $selectedXml = $xmlFiles[0]
} else {
    Write-Host "=== XML files found ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $xmlFiles.Count; $i++) {
        Write-Host ("  {0}: {1}" -f ($i + 1), $xmlFiles[$i].Name)
    }
    Write-Host ""
    $choice = Read-Host "Select number (1-$($xmlFiles.Count))"
    $idx = [int]$choice - 1
    if ($idx -lt 0 -or $idx -ge $xmlFiles.Count) {
        Write-Host "[ERROR] Invalid selection." -ForegroundColor Red
        Read-Host "Enter to exit"; exit 1
    }
    $selectedXml = $xmlFiles[$idx]
}

$InputXml   = $selectedXml.FullName
$OutputXlsx = Join-Path $ScriptDir ($selectedXml.BaseName + ".xlsx")

Write-Host ""
Write-Host "Input : $InputXml"
Write-Host "Output: $OutputXlsx"
Write-Host ""

if (-not (Test-Path -LiteralPath $InputXml)) {
    Write-Host "[ERROR] XML not found: $InputXml" -ForegroundColor Red
    Read-Host "Enter to exit"; exit 1
}

Write-Host "Loading XML..." -ForegroundColor Cyan
[xml]$xml = Get-Content -LiteralPath $InputXml -Encoding UTF8
$devices = $xml.preset.device
$sep = [char]0x00AC

$presetName = ""
$presetDesc = ""
$presetVer  = ""
if ($xml.preset.name)        { $presetName = $xml.preset.name }
if ($xml.preset.description) { $presetDesc = $xml.preset.description }
if ($xml.preset.version)     { $presetVer  = $xml.preset.version }

# === Device list ===
$deviceList = foreach ($dev in $devices) {
    $ipAddr0 = ""; $ipMode0 = ""; $netId0 = ""
    $ipAddr1 = ""; $ipMode1 = ""; $netId1 = ""
    $interfaces = @($dev.interface)
    foreach ($iface in $interfaces) {
        if ($null -eq $iface) { continue }
        $nid = ""
        if ($iface.network) { $nid = $iface.network }
        $ia = ""; $im = ""
        $ipv4 = $iface.ipv4_address
        if ($ipv4) {
            if ($ipv4.address)  { $ia = $ipv4.address }
            if ($ipv4.mode)     { $im = $ipv4.mode }
            if ($ipv4.'#text')  { $ia = $ipv4.'#text' }
        }
        if ($ipAddr0 -eq "" -and $ipMode0 -eq "") {
            $netId0 = $nid; $ipAddr0 = $ia; $ipMode0 = $im
        } else {
            $netId1 = $nid; $ipAddr1 = $ia; $ipMode1 = $im
        }
    }

    $prefMaster = ""; $switchVlan = ""; $devTypeStr = ""; $devType = ""
    $devId = ""; $procId = ""; $mfId = ""; $modelId = ""
    $encoding = ""; $redundancy = ""; $extWordClk = ""

    if ($dev.preferred_master)    { $prefMaster = $dev.preferred_master.value }
    if ($dev.switch_vlan)         { $switchVlan = $dev.switch_vlan.value }
    if ($dev.device_type_string)  { $devTypeStr = $dev.device_type_string }
    if ($dev.device_type)         { $devType    = $dev.device_type }
    if ($dev.encoding)            { $encoding   = $dev.encoding }
    if ($dev.manufacturer_id)     { $mfId       = $dev.manufacturer_id }
    if ($dev.model_id)            { $modelId    = $dev.model_id }
    if ($dev.redundancy)          { $redundancy = $dev.redundancy.value }
    if ($dev.external_word_clock) { $extWordClk = $dev.external_word_clock.value }
    if ($dev.instance_id) {
        if ($dev.instance_id.device_id)  { $devId  = $dev.instance_id.device_id }
        if ($dev.instance_id.process_id) { $procId = $dev.instance_id.process_id }
    }

    $interopMode = ""; $aes67Prefix = ""
    if ($dev.rtp) {
        if ($dev.rtp.interop_mode)                   { $interopMode = $dev.rtp.interop_mode }
        if ($dev.rtp.aes67_multicast_address_prefix) { $aes67Prefix = $dev.rtp.aes67_multicast_address_prefix }
    }

    $clockSubdomain = ""; $clockV1Enabled = ""; $clockV2Enabled = ""
    $clockV1Unicast = ""; $clockV2Unicast = ""
    if ($dev.clock) {
        if ($dev.clock.subdomain_name)            { $clockSubdomain = $dev.clock.subdomain_name }
        if ($dev.clock.v1_enabled)                { $clockV1Enabled = $dev.clock.v1_enabled }
        if ($dev.clock.v2_enabled)                { $clockV2Enabled = $dev.clock.v2_enabled }
        if ($dev.clock.v1_unicast_delay_requests) { $clockV1Unicast = $dev.clock.v1_unicast_delay_requests }
        if ($dev.clock.v2_unicast_delay_requests) { $clockV2Unicast = $dev.clock.v2_unicast_delay_requests }
    }

    $clockPreferred = ""; $clockFollowerOnly = ""
    if ($dev.clock_priority) {
        if ($dev.clock_priority.preferred)     { $clockPreferred    = $dev.clock_priority.preferred }
        if ($dev.clock_priority.follower_only) { $clockFollowerOnly = $dev.clock_priority.follower_only }
    }

    [PSCustomObject]@{
        DeviceName       = $dev.name
        DefaultName      = $dev.default_name
        FriendlyName     = $dev.friendly_name
        Model            = $dev.model_name
        Manufacturer     = $dev.manufacturer_name
        ManufacturerId   = $mfId
        ModelId          = $modelId
        ModelVersion     = $dev.model_version
        DeviceType       = $devType
        DeviceTypeString = $devTypeStr
        DeviceId         = $devId
        ProcessId        = $procId
        SampleRate       = $dev.samplerate
        Encoding         = $encoding
        Latency          = $dev.unicast_latency
        Redundancy       = $redundancy
        ExtWordClock     = $extWordClk
        PriNetwork       = $netId0
        PriIPv4Address   = $ipAddr0
        PriIPv4Mode      = $ipMode0
        SecNetwork       = $netId1
        SecIPv4Address   = $ipAddr1
        SecIPv4Mode      = $ipMode1
        SwitchVlan       = $switchVlan
        PreferredMaster  = $prefMaster
        InteropMode      = $interopMode
        AES67McPrefix    = $aes67Prefix
        ClockSubdomain   = $clockSubdomain
        ClockV1Enabled   = $clockV1Enabled
        ClockV2Enabled   = $clockV2Enabled
        ClockV1Unicast   = $clockV1Unicast
        ClockV2Unicast   = $clockV2Unicast
        ClockPreferred   = $clockPreferred
        ClockFollowerOnly = $clockFollowerOnly
        TxCount          = [string]($dev.txchannel | Measure-Object).Count
        RxCount          = [string]($dev.rxchannel | Measure-Object).Count
    }
}

# === TX list ===
$txList = [System.Collections.ArrayList]::new()
foreach ($dev in $devices) {
    foreach ($tx in $dev.txchannel) {
        $k = $dev.name + $sep + $tx.label
        [void]$txList.Add([PSCustomObject]@{
            Device    = $dev.name
            DanteId   = $tx.danteId
            Label     = $tx.label
            MediaType = $tx.mediaType
            Key       = $k
        })
    }
}

# === TX Flow list ===
$txFlowList = [System.Collections.ArrayList]::new()
foreach ($dev in $devices) {
    foreach ($flow in $dev.txflow) {
        if ($null -eq $flow) { continue }
        $destAddr = ""; $destPort = ""
        if ($flow.destinationAddress) {
            if ($flow.destinationAddress.address) { $destAddr = $flow.destinationAddress.address }
            if ($flow.destinationAddress.port)    { $destPort = $flow.destinationAddress.port }
        }
        $fpp = ""; $sid = ""; $ttype = ""
        if ($flow.fpp)           { $fpp   = $flow.fpp }
        if ($flow.sessionId)     { $sid   = $flow.sessionId }
        if ($flow.transportType) { $ttype = $flow.transportType }

        $slots = [System.Collections.ArrayList]::new()
        foreach ($slot in $flow.slot) {
            if ($null -eq $slot) { continue }
            [void]$slots.Add($slot.channelId)
        }

        $flowType = "Dante"
        if ($destAddr -ne "" -and $fpp -ne "" -and $ttype -eq "2") {
            $flowType = "AES67"
        }

        [void]$txFlowList.Add([PSCustomObject]@{
            Device        = $dev.name
            DanteId       = $flow.danteId
            FPP           = $fpp
            MediaType     = $flow.mediaType
            SessionId     = $sid
            TransportType = $ttype
            DestAddress   = $destAddr
            DestPort      = $destPort
            SlotChannels  = ($slots -join ", ")
            SlotCount     = [string]$slots.Count
            FlowType      = $flowType
        })
    }
}

# === RX list ===
$rxList = [System.Collections.ArrayList]::new()
foreach ($dev in $devices) {
    foreach ($rx in $dev.rxchannel) {
        $subCh  = $null
        $subDev = $null
        if ($rx.subscribed_channel) { $subCh  = $rx.subscribed_channel }
        if ($rx.subscribed_device)  { $subDev = $rx.subscribed_device }
        $k = $dev.name + $sep + $rx.name
        [void]$rxList.Add([PSCustomObject]@{
            Device            = $dev.name
            DanteId           = $rx.danteId
            Name              = $rx.name
            MediaType         = $rx.mediaType
            SubscribedChannel = $subCh
            SubscribedDevice  = $subDev
            Key               = $k
        })
    }
}

# === Subscription list ===
$subList = [System.Collections.ArrayList]::new()
foreach ($rx in $rxList) {
    if ($null -eq $rx.SubscribedDevice) { continue }
    if ($null -eq $rx.SubscribedChannel) { continue }
    [void]$subList.Add([PSCustomObject]@{
        RxDevice  = $rx.Device
        RxChannel = $rx.Name
        RxDanteId = $rx.DanteId
        TxDevice  = $rx.SubscribedDevice
        TxChannel = $rx.SubscribedChannel
        MediaType = $rx.MediaType
    })
}

$txKeyIndex = @{}
for ($t = 0; $t -lt $txList.Count; $t++) {
    $txKeyIndex[$txList[$t].Key] = $t
}

# ============================================================
# Excel
# ============================================================
Write-Host "Starting Excel..." -ForegroundColor Cyan
Write-Host "  TX: $($txList.Count) / RX: $($rxList.Count) channels" -ForegroundColor Gray
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$wb = $excel.Workbooks.Add()

function Set-CellText($ws, [int]$row, [int]$col, $value) {
    $cell = $ws.Cells.Item($row, $col)
    $cell.NumberFormat = "@"
    if ($null -ne $value) { $cell.Value2 = [string]$value }
}

function Write-Header($ws, [string[]]$headers, [int]$row) {
    for ($c = 0; $c -lt $headers.Count; $c++) {
        Set-CellText $ws $row ($c + 1) $headers[$c]
    }
    $rng = $ws.Range($ws.Cells.Item($row, 1), $ws.Cells.Item($row, $headers.Count))
    $rng.Font.Bold = $true
    $rng.Interior.Color = 0xD9E1F2
}

# ==========================================================
# Devices ヘッダー/プロパティ
# ==========================================================
if ($detailMode) {
    $devH = @(
        "Device Name", "Default Name", "Friendly Name",
        "Model", "Manufacturer", "Manufacturer ID",
        "Model ID", "Model Version",
        "Device Type", "Device Type String",
        "Device ID", "Process ID",
        "Sample Rate", "Encoding", "Latency (us)",
        "Redundancy", "External Word Clock",
        "Pri Network", "Pri IPv4 Address", "Pri IPv4 Mode",
        "Sec Network", "Sec IPv4 Address", "Sec IPv4 Mode",
        "Switch VLAN", "Preferred Master",
        "Interop Mode", "AES67 MC Prefix",
        "Clock Subdomain",
        "PTP v1 Enabled", "PTP v2 Enabled",
        "PTP v1 Unicast Delay", "PTP v2 Unicast Delay",
        "Clock Preferred", "Clock Follower Only",
        "TX Ch", "RX Ch"
    )
    $devProps = @(
        "DeviceName", "DefaultName", "FriendlyName",
        "Model", "Manufacturer", "ManufacturerId",
        "ModelId", "ModelVersion",
        "DeviceType", "DeviceTypeString",
        "DeviceId", "ProcessId",
        "SampleRate", "Encoding", "Latency",
        "Redundancy", "ExtWordClock",
        "PriNetwork", "PriIPv4Address", "PriIPv4Mode",
        "SecNetwork", "SecIPv4Address", "SecIPv4Mode",
        "SwitchVlan", "PreferredMaster",
        "InteropMode", "AES67McPrefix",
        "ClockSubdomain",
        "ClockV1Enabled", "ClockV2Enabled",
        "ClockV1Unicast", "ClockV2Unicast",
        "ClockPreferred", "ClockFollowerOnly",
        "TxCount", "RxCount"
    )
} else {
    $devH = @(
        "Device Name", "Default Name", "Friendly Name",
        "Model", "Manufacturer", "Model Version",
        "Device Type String",
        "Sample Rate", "Encoding", "Latency (us)",
        "Redundancy", "External Word Clock",
        "Pri IPv4 Address", "Pri IPv4 Mode",
        "Sec IPv4 Address", "Sec IPv4 Mode",
        "Preferred Master", "Interop Mode", "Clock Preferred"
    )
    $devProps = @(
        "DeviceName", "DefaultName", "FriendlyName",
        "Model", "Manufacturer", "ModelVersion",
        "DeviceTypeString",
        "SampleRate", "Encoding", "Latency",
        "Redundancy", "ExtWordClock",
        "PriIPv4Address", "PriIPv4Mode",
        "SecIPv4Address", "SecIPv4Mode",
        "PreferredMaster", "InteropMode", "ClockPreferred"
    )
}

# ==========================================================
# Sheet1: Devices
# ==========================================================
Write-Host "  Sheet1: Devices..." -ForegroundColor Gray
$ws1 = $wb.Worksheets.Item(1)
$ws1.Name = "Devices"
$ws1.Cells.NumberFormat = "@"

Set-CellText $ws1 1 1 "Preset Name"
Set-CellText $ws1 1 2 $presetName
Set-CellText $ws1 2 1 "Description"
Set-CellText $ws1 2 2 $presetDesc
Set-CellText $ws1 3 1 "Preset Version"
Set-CellText $ws1 3 2 $presetVer
$ws1.Range("A1:A3").Font.Bold = $true

$startRow = 5
Write-Header $ws1 $devH $startRow

$r = $startRow + 1
foreach ($d in $deviceList) {
    for ($c = 0; $c -lt $devProps.Count; $c++) {
        Set-CellText $ws1 $r ($c + 1) $d.($devProps[$c])
    }
    $r++
}
$ws1.Range($ws1.Columns.Item(1), $ws1.Columns.Item($devH.Count)).AutoFit() | Out-Null

# ==========================================================
# Sheet2: Patch Matrix
# ==========================================================
Write-Host "  Sheet2: Patch Matrix..." -ForegroundColor Gray
$ws2 = $wb.Worksheets.Add([System.Type]::Missing, $ws1)
$ws2.Name = "Patch Matrix"
$oR = 3; $oC = 3

if ($txList.Count -gt 0 -and $rxList.Count -gt 0) {
    Write-Host "    Matrix: $($rxList.Count) rows x $($txList.Count) cols" -ForegroundColor Gray

    $curDev = ""; $startCol = $oC
    for ($t = 0; $t -lt $txList.Count; $t++) {
        $col = $oC + $t
        $tx = $txList[$t]
        Set-CellText $ws2 2 $col $tx.Label
        if ($tx.Device -ne $curDev) {
            if (($curDev -ne "") -and (($col - 1) -gt $startCol)) {
                $mr = $ws2.Range($ws2.Cells.Item(1, $startCol), $ws2.Cells.Item(1, $col - 1))
                $mr.Merge() | Out-Null
            }
            Set-CellText $ws2 1 $col $tx.Device
            $curDev = $tx.Device; $startCol = $col
        }
    }
    $lastTxCol = $oC + $txList.Count - 1
    if ($lastTxCol -gt $startCol) {
        $mr = $ws2.Range($ws2.Cells.Item(1, $startCol), $ws2.Cells.Item(1, $lastTxCol))
        $mr.Merge() | Out-Null
    }

    $curDev = ""; $startRow = $oR
    for ($i = 0; $i -lt $rxList.Count; $i++) {
        $row = $oR + $i
        $rx = $rxList[$i]
        Set-CellText $ws2 $row 2 $rx.Name
        if ($rx.Device -ne $curDev) {
            if (($curDev -ne "") -and (($row - 1) -gt $startRow)) {
                $mr = $ws2.Range($ws2.Cells.Item($startRow, 1), $ws2.Cells.Item($row - 1, 1))
                $mr.Merge() | Out-Null
            }
            Set-CellText $ws2 $row 1 $rx.Device
            $curDev = $rx.Device; $startRow = $row
        }
    }
    $lastRxRow = $oR + $rxList.Count - 1
    if ($lastRxRow -gt $startRow) {
        $mr = $ws2.Range($ws2.Cells.Item($startRow, 1), $ws2.Cells.Item($lastRxRow, 1))
        $mr.Merge() | Out-Null
    }

    Write-Host "    Marking cross points..." -ForegroundColor Gray
    for ($i = 0; $i -lt $rxList.Count; $i++) {
        $rx = $rxList[$i]
        if ($null -eq $rx.SubscribedDevice) { continue }
        if ($null -eq $rx.SubscribedChannel) { continue }
        $txKey = $rx.SubscribedDevice + $sep + $rx.SubscribedChannel
        if ($txKeyIndex.ContainsKey($txKey)) {
            $t = $txKeyIndex[$txKey]
            $cR = $oR + $i; $cC = $oC + $t
            $ws2.Cells.Item($cR, $cC).Interior.Color = 0xC6EFCE
        }
    }

    Write-Host "    Formatting..." -ForegroundColor Gray
    $txH1 = $ws2.Range($ws2.Cells.Item(1, $oC), $ws2.Cells.Item(1, $lastTxCol))
    $txH1.Font.Bold = $true; $txH1.HorizontalAlignment = -4108
    $txH1.Interior.Color = 0xFFD966; $txH1.Font.Size = 8
    $txH1.Orientation = 90

    $maxDevLen = 0
    foreach ($tx in $txList) {
        if ($tx.Device.Length -gt $maxDevLen) { $maxDevLen = $tx.Device.Length }
    }
    $row1Height = [Math]::Max(50, $maxDevLen * 7)
    $ws2.Rows.Item(1).RowHeight = $row1Height

    $txH2 = $ws2.Range($ws2.Cells.Item(2, $oC), $ws2.Cells.Item(2, $lastTxCol))
    $txH2.Font.Bold = $true; $txH2.HorizontalAlignment = -4108
    $txH2.Interior.Color = 0xFFF2CC; $txH2.Orientation = 90; $txH2.Font.Size = 7

    $rxH1 = $ws2.Range($ws2.Cells.Item($oR, 1), $ws2.Cells.Item($lastRxRow, 1))
    $rxH1.Font.Bold = $true; $rxH1.VerticalAlignment = -4108
    $rxH1.Interior.Color = 0x9BC2E6; $rxH1.Font.Size = 8

    $rxH2 = $ws2.Range($ws2.Cells.Item($oR, 2), $ws2.Cells.Item($lastRxRow, 2))
    $rxH2.Font.Bold = $true; $rxH2.Interior.Color = 0xDDEBF7; $rxH2.Font.Size = 7

    Set-CellText $ws2 1 1 "RX / TX"
    $ws2.Cells.Item(1, 1).Font.Bold = $true
    $matR = $ws2.Range($ws2.Cells.Item(1, 1), $ws2.Cells.Item($lastRxRow, $lastTxCol))
    $matR.Borders.LineStyle = 1; $matR.Borders.Weight = 2; $matR.Borders.Color = 0xD0D0D0

    $ws2.Columns.Item(1).ColumnWidth = 18
    $ws2.Columns.Item(2).ColumnWidth = 16
    $txColRange = $ws2.Range($ws2.Cells.Item(1, $oC), $ws2.Cells.Item(1, $lastTxCol)).EntireColumn
    $txColRange.ColumnWidth = 3
    $rxRowRange = $ws2.Range($ws2.Cells.Item($oR, 1), $ws2.Cells.Item($lastRxRow, 1)).EntireRow
    $rxRowRange.RowHeight = 13

    $ws2.Activate()
    $ws2.Cells.Item($oR, $oC).Select()
    $excel.ActiveWindow.FreezePanes = $true
} else {
    Set-CellText $ws2 1 1 "(No TX/RX channels to display)"
}

# ==========================================================
# Sheet3: TX Flows
# ==========================================================
Write-Host "  Sheet3: TX Flows..." -ForegroundColor Gray
$ws3 = $wb.Worksheets.Add([System.Type]::Missing, $ws2)
$ws3.Name = "TX Flows"
$ws3.Cells.NumberFormat = "@"

if ($detailMode) {
    $flowH = @("Device Name", "Flow Type", "Dante ID", "FPP", "Media Type", "Session ID", "Transport Type", "Dest Address", "Dest Port", "Slot Count", "Slot Channels")
} else {
    $flowH = @("Device Name", "Flow Type", "Dest Address", "Dest Port", "Slot Count", "Slot Channels")
}
Write-Header $ws3 $flowH 1

$r = 2
foreach ($f in $txFlowList) {
    if ($detailMode) {
        Set-CellText $ws3 $r 1  $f.Device
        Set-CellText $ws3 $r 2  $f.FlowType
        Set-CellText $ws3 $r 3  $f.DanteId
        Set-CellText $ws3 $r 4  $f.FPP
        Set-CellText $ws3 $r 5  $f.MediaType
        Set-CellText $ws3 $r 6  $f.SessionId
        Set-CellText $ws3 $r 7  $f.TransportType
        Set-CellText $ws3 $r 8  $f.DestAddress
        Set-CellText $ws3 $r 9  $f.DestPort
        Set-CellText $ws3 $r 10 $f.SlotCount
        Set-CellText $ws3 $r 11 $f.SlotChannels
    } else {
        Set-CellText $ws3 $r 1 $f.Device
        Set-CellText $ws3 $r 2 $f.FlowType
        Set-CellText $ws3 $r 3 $f.DestAddress
        Set-CellText $ws3 $r 4 $f.DestPort
        Set-CellText $ws3 $r 5 $f.SlotCount
        Set-CellText $ws3 $r 6 $f.SlotChannels
    }
    $r++
}
if ($txFlowList.Count -eq 0) { Set-CellText $ws3 2 1 "(No TX Flows)" }

$lastFlowCol = if ($detailMode) { "K" } else { "F" }
$ws3.Columns.Item("A:$lastFlowCol").AutoFit() | Out-Null

$lastSheet = $ws3

# ==========================================================
# 詳細モードのみ
# ==========================================================
if ($detailMode) {
    Write-Host "  Sheet4: TX Channels..." -ForegroundColor Gray
    $ws4 = $wb.Worksheets.Add([System.Type]::Missing, $lastSheet)
    $ws4.Name = "TX Channels"
    $ws4.Cells.NumberFormat = "@"
    Write-Header $ws4 @("Device Name", "Dante ID", "Channel Label", "Media Type") 1
    $r = 2
    foreach ($tx in $txList) {
        Set-CellText $ws4 $r 1 $tx.Device
        Set-CellText $ws4 $r 2 $tx.DanteId
        Set-CellText $ws4 $r 3 $tx.Label
        Set-CellText $ws4 $r 4 $tx.MediaType
        $r++
    }
    $ws4.Columns.Item("A:D").AutoFit() | Out-Null
    $lastSheet = $ws4

    Write-Host "  Sheet5: RX Channels..." -ForegroundColor Gray
    $ws5 = $wb.Worksheets.Add([System.Type]::Missing, $lastSheet)
    $ws5.Name = "RX Channels"
    $ws5.Cells.NumberFormat = "@"
    Write-Header $ws5 @("Device Name", "Dante ID", "Channel Name", "Media Type", "Subscribed Device", "Subscribed Channel", "Status") 1
    $r = 2
    foreach ($rx in $rxList) {
        Set-CellText $ws5 $r 1 $rx.Device
        Set-CellText $ws5 $r 2 $rx.DanteId
        Set-CellText $ws5 $r 3 $rx.Name
        Set-CellText $ws5 $r 4 $rx.MediaType
        $status = "Unsubscribed"
        if ($null -ne $rx.SubscribedDevice -and $null -ne $rx.SubscribedChannel) {
            Set-CellText $ws5 $r 5 $rx.SubscribedDevice
            Set-CellText $ws5 $r 6 $rx.SubscribedChannel
            $status = "Connected"
        }
        Set-CellText $ws5 $r 7 $status
        $r++
    }
    $ws5.Columns.Item("A:G").AutoFit() | Out-Null
    $lastSheet = $ws5

    Write-Host "  Sheet6: Subscriptions..." -ForegroundColor Gray
    $ws6 = $wb.Worksheets.Add([System.Type]::Missing, $lastSheet)
    $ws6.Name = "Subscriptions"
    $ws6.Cells.NumberFormat = "@"
    Write-Header $ws6 @("No.", "RX Device", "RX Channel", "RX Dante ID", "TX Device", "TX Channel", "Media Type") 1
    $r = 2; $no = 1
    foreach ($s in $subList) {
        Set-CellText $ws6 $r 1 $no
        Set-CellText $ws6 $r 2 $s.RxDevice
        Set-CellText $ws6 $r 3 $s.RxChannel
        Set-CellText $ws6 $r 4 $s.RxDanteId
        Set-CellText $ws6 $r 5 $s.TxDevice
        Set-CellText $ws6 $r 6 $s.TxChannel
        Set-CellText $ws6 $r 7 $s.MediaType
        $r++; $no++
    }
    if ($subList.Count -eq 0) { Set-CellText $ws6 2 1 "(No subscriptions)" }
    $ws6.Columns.Item("A:G").AutoFit() | Out-Null
}

# ==========================================================
# Save
# ==========================================================
Write-Host "Saving..." -ForegroundColor Cyan
$excel.ScreenUpdating = $true
$ws1.Activate()
if (Test-Path -LiteralPath $OutputXlsx) { Remove-Item -LiteralPath $OutputXlsx -Force }
$wb.SaveAs($OutputXlsx, 51)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()

$modeLabel = if ($detailMode) { "Detail" } else { "Default" }
Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host " Done: $OutputXlsx" -ForegroundColor Green
Write-Host " Mode:          $modeLabel" -ForegroundColor Cyan
Write-Host " Devices:       $($devices.Count)" -ForegroundColor Cyan
Write-Host " TX Channels:   $($txList.Count)" -ForegroundColor Cyan
Write-Host " TX Flows:      $($txFlowList.Count)" -ForegroundColor Cyan
Write-Host " RX Channels:   $($rxList.Count)" -ForegroundColor Cyan
Write-Host " Subscriptions: $($subList.Count)" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Green
Read-Host "Enter to exit"