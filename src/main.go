package main

import (
	"bufio"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

const helpText = `================================================================================
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
  - Microsoft Excel NOT required (Go native Excel generator)


  HOW TO USE
  ----------
  1. Place this tool in any folder.
  2. Place the Dante preset XML file(s) in the same folder.
  3. Run the executable.
  4. A menu appears:
       === Menu ===
         1: Default (summary)
         2: Detail  (all info)
         H: Help
  5. Select mode (1 / 2 / H).
  6. An .xlsx file is created in the same folder.
================================================================================`

// XML structures
type Preset struct {
	XMLName     xml.Name `xml:"preset"`
	Name        string   `xml:"name"`
	Description string   `xml:"description"`
	Version     string   `xml:"version"`
	VersionAttr string   `xml:"version,attr"`
	Devices     []Device `xml:"device"`
}

func (p *Preset) GetVersion() string {
	if p.VersionAttr != "" {
		return p.VersionAttr
	}
	return p.Version
}

type ValueElement struct {
	Value string `xml:"value,attr"`
}

type Device struct {
	Name              string        `xml:"name"`
	DefaultName       string        `xml:"default_name"`
	FriendlyName      string        `xml:"friendly_name"`
	ModelName         string        `xml:"model_name"`
	ManufacturerName  string        `xml:"manufacturer_name"`
	ModelVersion      string        `xml:"model_version"`
	SampleRate        string        `xml:"samplerate"`
	Encoding          string        `xml:"encoding"`
	PullUpValue       string        `xml:"pull_up_value"`
	UnicastLatency    string        `xml:"unicast_latency"`
	Interfaces        []Interface   `xml:"interface"`
	Clock             Clock         `xml:"clock"`
	ClockPriority     ClockPriority `xml:"clock_priority"`
	Rtp               Rtp           `xml:"rtp"`
	TxChannels        []TxChannel   `xml:"txchannel"`
	RxChannels        []RxChannel   `xml:"rxchannel"`
	TxFlows           []TxFlow      `xml:"txflow"`
	PreferredMaster   ValueElement  `xml:"preferred_master"`
	SwitchVlan        ValueElement  `xml:"switch_vlan"`
	Redundancy        ValueElement  `xml:"redundancy"`
	ExternalWordClock ValueElement  `xml:"external_word_clock"`
	ManufacturerId    string        `xml:"manufacturer_id"`
	ModelId           string        `xml:"model_id"`
	DeviceType        string        `xml:"device_type"`
	DeviceTypeString  string        `xml:"device_type_string"`
	InstanceId        InstanceId    `xml:"instance_id"`
}

type InstanceId struct {
	DeviceId  string `xml:"device_id"`
	ProcessId string `xml:"process_id"`
}

type Interface struct {
	Network     string      `xml:"network,attr"`
	IPv4Address IPv4Address `xml:"ipv4_address"`
}

type IPv4Address struct {
	Mode    string `xml:"mode,attr"`
	Address string `xml:"address,attr"`
	Netmask string `xml:"netmask,attr"` // New
	Gateway string `xml:"gateway,attr"` // New
	Value   string `xml:",chardata"`
}

func (ip *IPv4Address) GetAddress() string {
	if ip.Address != "" {
		return ip.Address
	}
	return strings.TrimSpace(ip.Value)
}

type Clock struct {
	SubdomainName          string `xml:"subdomain_name"`
	V1Enabled              string `xml:"v1_enabled"`
	V2Enabled              string `xml:"v2_enabled"`
	V1UnicastDelayRequests string `xml:"v1_unicast_delay_requests"`
	V2UnicastDelayRequests string `xml:"v2_unicast_delay_requests"`
	V2DomainNumber         string `xml:"v2_domain_number"` // New
	V2DSCP                 string `xml:"v2_dscp"`          // New
	MulticastTTL           string `xml:"multicast_ttl"`    // New
}

type ClockPriority struct {
	Preferred    string `xml:"preferred"`
	FollowerOnly string `xml:"follower_only"`
	V2Priority1  string `xml:"v2_priority1"` // New
	V2Priority2  string `xml:"v2_priority2"` // New
}

type Rtp struct {
	InteropMode                 string `xml:"interop_mode"`
	Aes67MulticastAddressPrefix string `xml:"aes67_multicast_address_prefix"`
}

type TxChannel struct {
	DanteId   string `xml:"danteId,attr"`
	MediaType string `xml:"mediaType,attr"`
	Label     string `xml:"label"`
}

type RxChannel struct {
	DanteId           string `xml:"danteId,attr"`
	MediaType         string `xml:"mediaType,attr"`
	Name              string `xml:"name"`
	SubscribedChannel string `xml:"subscribed_channel"`
	SubscribedDevice  string `xml:"subscribed_device"`
}

type TxFlow struct {
	DanteId            string             `xml:"danteId,attr"`
	Fpp                string             `xml:"fpp,attr"`
	MediaType          string             `xml:"mediaType,attr"`
	SessionId          string             `xml:"sessionId,attr"`
	TransportType      string             `xml:"transportType,attr"`
	Slots              []Slot             `xml:"slot"`
	DestinationAddress DestinationAddress `xml:"destinationAddress"`
}

type Slot struct {
	ChannelId string `xml:"channelId,attr"`
}

type DestinationAddress struct {
	Address string `xml:"address,attr"`
	Port    string `xml:"port,attr"`
}

// Flat structured rows for Excel mapping
type DeviceRow struct {
	DeviceName        string
	DefaultName       string
	FriendlyName      string
	Model             string
	Manufacturer      string
	ManufacturerId    string
	ModelId           string
	ModelVersion      string
	DeviceType        string
	DeviceTypeString  string
	DeviceId          string
	ProcessId         string
	SampleRate        string
	Encoding          string
	Latency           string
	Redundancy        string
	ExtWordClock      string
	PriNetwork        string
	PriIPv4Address    string
	PriIPv4Mode       string
	PriIPv4Netmask    string // New
	PriIPv4Gateway    string // New
	SecNetwork        string
	SecIPv4Address    string
	SecIPv4Mode       string
	SecIPv4Netmask    string // New
	SecIPv4Gateway    string // New
	SwitchVlan        string
	PreferredMaster   string
	InteropMode       string
	AES67McPrefix     string
	ClockSubdomain    string
	ClockV1Enabled    string
	ClockV2Enabled    string
	ClockV2Domain     string // New
	ClockV2DSCP       string // New
	ClockMulticastTTL string // New
	ClockV1Unicast    string
	ClockV2Unicast    string
	ClockPreferred    string
	ClockFollowerOnly string
	ClockV2Priority1  string // New
	ClockV2Priority2  string // New
	PullUpValue       string // New
	TxCount           int
	RxCount           int
}

type TxRow struct {
	Device    string
	DanteId   string
	Label     string
	MediaType string
	Key       string // Format: DeviceName + "\xac" + Label
}

type TxFlowRow struct {
	Device        string
	DanteId       string
	Fpp           string
	MediaType     string
	SessionId     string
	TransportType string
	DestAddress   string
	DestPort      string
	SlotChannels  string // Comma joined slot channel IDs
	SlotCount     int
	FlowType      string // "AES67" or "Dante"
}

type RxRow struct {
	Device            string
	DanteId           string
	Name              string
	MediaType         string
	SubscribedChannel string
	SubscribedDevice  string
	Key               string // Format: DeviceName + "\xac" + ChannelName
}

type SubRow struct {
	RxDevice  string
	RxChannel string
	RxDanteId string
	TxDevice  string
	TxChannel string
	MediaType string
}

const Sep = "\u00ac" // Same separator as PowerShell [char]0x00AC

func ProcessPresetData(preset *Preset) ([]DeviceRow, []TxRow, []TxFlowRow, []RxRow, []SubRow) {
	var deviceRows []DeviceRow
	var txRows []TxRow
	var txFlowRows []TxFlowRow
	var rxRows []RxRow
	var subRows []SubRow

	// Device List Mapping
	for _, dev := range preset.Devices {
		var ipAddr0, ipMode0, netId0, netmask0, gateway0 string
		var ipAddr1, ipMode1, netId1, netmask1, gateway1 string

		for _, iface := range dev.Interfaces {
			nid := iface.Network
			ia := iface.IPv4Address.GetAddress()
			im := iface.IPv4Address.Mode
			nm := iface.IPv4Address.Netmask
			gw := iface.IPv4Address.Gateway

			if ipAddr0 == "" && ipMode0 == "" {
				netId0 = nid
				ipAddr0 = ia
				ipMode0 = im
				netmask0 = nm
				gateway0 = gw
			} else if ipAddr1 == "" && ipMode1 == "" {
				netId1 = nid
				ipAddr1 = ia
				ipMode1 = im
				netmask1 = nm
				gateway1 = gw
			}
		}

		// Count tx & rx channels
		txCount := len(dev.TxChannels)
		rxCount := len(dev.RxChannels)

		devRow := DeviceRow{
			DeviceName:        dev.Name,
			DefaultName:       dev.DefaultName,
			FriendlyName:      dev.FriendlyName,
			Model:             dev.ModelName,
			Manufacturer:      dev.ManufacturerName,
			ManufacturerId:    dev.ManufacturerId,
			ModelId:           dev.ModelId,
			ModelVersion:      dev.ModelVersion,
			DeviceType:        dev.DeviceType,
			DeviceTypeString:  dev.DeviceTypeString,
			DeviceId:          dev.InstanceId.DeviceId,
			ProcessId:         dev.InstanceId.ProcessId,
			SampleRate:        dev.SampleRate,
			Encoding:          dev.Encoding,
			Latency:           dev.UnicastLatency,
			Redundancy:        dev.Redundancy.Value,
			ExtWordClock:      dev.ExternalWordClock.Value,
			PriNetwork:        netId0,
			PriIPv4Address:    ipAddr0,
			PriIPv4Mode:       ipMode0,
			PriIPv4Netmask:    netmask0,
			PriIPv4Gateway:    gateway0,
			SecNetwork:        netId1,
			SecIPv4Address:    ipAddr1,
			SecIPv4Mode:       ipMode1,
			SecIPv4Netmask:    netmask1,
			SecIPv4Gateway:    gateway1,
			SwitchVlan:        dev.SwitchVlan.Value,
			PreferredMaster:   dev.PreferredMaster.Value,
			InteropMode:       dev.Rtp.InteropMode,
			AES67McPrefix:     dev.Rtp.Aes67MulticastAddressPrefix,
			ClockSubdomain:    dev.Clock.SubdomainName,
			ClockV1Enabled:    dev.Clock.V1Enabled,
			ClockV2Enabled:    dev.Clock.V2Enabled,
			ClockV2Domain:     dev.Clock.V2DomainNumber,
			ClockV2DSCP:       dev.Clock.V2DSCP,
			ClockMulticastTTL: dev.Clock.MulticastTTL,
			ClockV1Unicast:    dev.Clock.V1UnicastDelayRequests,
			ClockV2Unicast:    dev.Clock.V2UnicastDelayRequests,
			ClockPreferred:    dev.ClockPriority.Preferred,
			ClockFollowerOnly: dev.ClockPriority.FollowerOnly,
			ClockV2Priority1:  dev.ClockPriority.V2Priority1,
			ClockV2Priority2:  dev.ClockPriority.V2Priority2,
			PullUpValue:       dev.PullUpValue,
			TxCount:           txCount,
			RxCount:           rxCount,
		}

		// Tx Channels
		for _, tx := range dev.TxChannels {
			txRows = append(txRows, TxRow{
				Device:    dev.Name,
				DanteId:   tx.DanteId,
				Label:     tx.Label,
				MediaType: tx.MediaType,
				Key:       dev.Name + Sep + tx.Label,
			})
		}

		// Tx Flows
		for _, flow := range dev.TxFlows {
			destAddr := flow.DestinationAddress.Address
			destPort := flow.DestinationAddress.Port

			var slots []string
			for _, slot := range flow.Slots {
				slots = append(slots, slot.ChannelId)
			}

			flowType := "Dante"
			if destAddr != "" && flow.Fpp != "" && flow.TransportType == "2" {
				flowType = "AES67"
			}

			txFlowRows = append(txFlowRows, TxFlowRow{
				Device:        dev.Name,
				DanteId:       flow.DanteId,
				Fpp:           flow.Fpp,
				MediaType:     flow.MediaType,
				SessionId:     flow.SessionId,
				TransportType: flow.TransportType,
				DestAddress:   destAddr,
				DestPort:      destPort,
				SlotChannels:  strings.Join(slots, ", "),
				SlotCount:     len(slots),
				FlowType:      flowType,
			})
		}

		// Rx Channels
		for _, rx := range dev.RxChannels {
			rxRows = append(rxRows, RxRow{
				Device:            dev.Name,
				DanteId:           rx.DanteId,
				Name:              rx.Name,
				MediaType:         rx.MediaType,
				SubscribedChannel: rx.SubscribedChannel,
				SubscribedDevice:  rx.SubscribedDevice,
				Key:               dev.Name + Sep + rx.Name,
			})
		}

		deviceRows = append(deviceRows, devRow)
	}

	// Subscriptions
	for _, rx := range rxRows {
		if rx.SubscribedDevice != "" && rx.SubscribedChannel != "" {
			subRows = append(subRows, SubRow{
				RxDevice:  rx.Device,
				RxChannel: rx.Name,
				RxDanteId: rx.DanteId,
				TxDevice:  rx.SubscribedDevice,
				TxChannel: rx.SubscribedChannel,
				MediaType: rx.MediaType,
			})
		}
	}

	return deviceRows, txRows, txFlowRows, rxRows, subRows
}

func main() {
	fmt.Println("========================================")
	fmt.Println("  Dante Preset XML -> Excel Converter")
	fmt.Println("========================================")
	fmt.Println("")

	// 1. XML File selection
	selectedXML, err := selectXMLFile()
	if err != nil {
		fmt.Printf("[ERROR] %v\n", err)
		waitForEnter()
		os.Exit(1)
	}

	// 2. Mode selection
	detailMode, err := selectMode()
	if err != nil {
		fmt.Printf("[ERROR] %v\n", err)
		waitForEnter()
		os.Exit(1)
	}

	baseName := strings.TrimSuffix(selectedXML, filepath.Ext(selectedXML))
	outputXlsx := baseName + ".xlsx"

	fmt.Println("")
	fmt.Printf("Input : %s\n", selectedXML)
	fmt.Printf("Output: %s\n", outputXlsx)
	fmt.Println("")

	fmt.Println("Loading XML...")
	xmlFile, err := os.Open(selectedXML)
	if err != nil {
		fmt.Printf("[ERROR] Failed to open XML: %v\n", err)
		waitForEnter()
		os.Exit(1)
	}
	defer xmlFile.Close()

	byteValue, err := io.ReadAll(xmlFile)
	if err != nil {
		fmt.Printf("[ERROR] Failed to read XML: %v\n", err)
		waitForEnter()
		os.Exit(1)
	}

	var preset Preset
	err = xml.Unmarshal(byteValue, &preset)
	if err != nil {
		fmt.Printf("[ERROR] Failed to parse XML: %v\n", err)
		waitForEnter()
		os.Exit(1)
	}

	deviceRows, txRows, txFlowRows, rxRows, subRows := ProcessPresetData(&preset)

	fmt.Printf("Parsed %d devices, %d TX channels, %d TX flows, %d RX channels, %d subscriptions.\n",
		len(preset.Devices), len(txRows), len(txFlowRows), len(rxRows), len(subRows))

	fmt.Println("Generating Excel...")
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("[ERROR] Failed to close file: %v\n", err)
		}
	}()

	// General Style Definitions
	textStyleID, err := f.NewStyle(&excelize.Style{
		NumFmt: 49, // '@' (text format)
	})
	if err != nil {
		fmt.Printf("[ERROR] Failed to create text style: %v\n", err)
	}

	headerStyleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"D9E1F2"}, Pattern: 1},
		Border: []excelize.Border{
			{Type: "top", Color: "D0D0D0", Style: 1},
			{Type: "bottom", Color: "D0D0D0", Style: 1},
			{Type: "left", Color: "D0D0D0", Style: 1},
			{Type: "right", Color: "D0D0D0", Style: 1},
		},
	})
	if err != nil {
		fmt.Printf("[ERROR] Failed to create header style: %v\n", err)
	}

	borderStyleID, err := f.NewStyle(&excelize.Style{
		NumFmt: 49,
		Border: []excelize.Border{
			{Type: "top", Color: "D0D0D0", Style: 1},
			{Type: "bottom", Color: "D0D0D0", Style: 1},
			{Type: "left", Color: "D0D0D0", Style: 1},
			{Type: "right", Color: "D0D0D0", Style: 1},
		},
	})
	if err != nil {
		fmt.Printf("[ERROR] Failed to create border style: %v\n", err)
	}

	// -------------------------------------------------------------
	// SHEET 1: Devices
	// -------------------------------------------------------------
	fmt.Println("  Writing sheet: Devices...")
	sheet1 := "Devices"
	f.SetSheetName("Sheet1", sheet1)
	f.SetColStyle(sheet1, "A:ZZ", textStyleID)

	// Preset basic metadata
	f.SetCellValue(sheet1, "A1", "Preset Name")
	f.SetCellValue(sheet1, "B1", preset.Name)
	f.SetCellValue(sheet1, "A2", "Description")
	f.SetCellValue(sheet1, "B2", preset.Description)
	f.SetCellValue(sheet1, "A3", "Preset Version")
	f.SetCellValue(sheet1, "B3", preset.GetVersion())

	// Bold for basic metadata labels
	boldStyleID, _ := f.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	f.SetCellStyle(sheet1, "A1", "A3", boldStyleID)

	var devH []string
	var devValuesFunc func(d DeviceRow) []string

	if detailMode {
		devH = []string{
			"Device Name", "Default Name", "Friendly Name",
			"Model", "Manufacturer", "Manufacturer ID",
			"Model ID", "Model Version", "Pull Up Value",
			"Device Type", "Device Type String",
			"Device ID", "Process ID",
			"Sample Rate", "Encoding", "Latency (us)",
			"Redundancy", "External Word Clock",
			"Pri Network", "Pri IPv4 Address", "Pri IPv4 Mode", "Pri IPv4 Netmask", "Pri IPv4 Gateway",
			"Sec Network", "Sec IPv4 Address", "Sec IPv4 Mode", "Sec IPv4 Netmask", "Sec IPv4 Gateway",
			"Switch VLAN", "Preferred Master",
			"Interop Mode", "AES67 MC Prefix",
			"Clock Subdomain",
			"PTP v1 Enabled", "PTP v2 Enabled",
			"PTP v2 Domain", "PTP v2 DSCP", "PTP v2 Multicast TTL",
			"PTP v1 Unicast Delay", "PTP v2 Unicast Delay",
			"Clock Preferred", "Clock Follower Only",
			"PTP v2 Priority 1", "PTP v2 Priority 2",
			"TX Ch", "RX Ch",
		}
		devValuesFunc = func(d DeviceRow) []string {
			return []string{
				d.DeviceName, d.DefaultName, d.FriendlyName,
				d.Model, d.Manufacturer, d.ManufacturerId,
				d.ModelId, d.ModelVersion, d.PullUpValue,
				d.DeviceType, d.DeviceTypeString,
				d.DeviceId, d.ProcessId,
				d.SampleRate, d.Encoding, d.Latency,
				d.Redundancy, d.ExtWordClock,
				d.PriNetwork, d.PriIPv4Address, d.PriIPv4Mode, d.PriIPv4Netmask, d.PriIPv4Gateway,
				d.SecNetwork, d.SecIPv4Address, d.SecIPv4Mode, d.SecIPv4Netmask, d.SecIPv4Gateway,
				d.SwitchVlan, d.PreferredMaster,
				d.InteropMode, d.AES67McPrefix,
				d.ClockSubdomain,
				d.ClockV1Enabled, d.ClockV2Enabled,
				d.ClockV2Domain, d.ClockV2DSCP, d.ClockMulticastTTL,
				d.ClockV1Unicast, d.ClockV2Unicast,
				d.ClockPreferred, d.ClockFollowerOnly,
				d.ClockV2Priority1, d.ClockV2Priority2,
				strconv.Itoa(d.TxCount), strconv.Itoa(d.RxCount),
			}
		}
	} else {
		devH = []string{
			"Device Name", "Default Name", "Friendly Name",
			"Model", "Manufacturer", "Model Version",
			"Device Type String",
			"Sample Rate", "Encoding", "Latency (us)",
			"Redundancy", "External Word Clock",
			"Pri IPv4 Address", "Pri IPv4 Mode",
			"Sec IPv4 Address", "Sec IPv4 Mode",
			"Preferred Master", "Interop Mode", "Clock Preferred",
		}
		devValuesFunc = func(d DeviceRow) []string {
			return []string{
				d.DeviceName, d.DefaultName, d.FriendlyName,
				d.Model, d.Manufacturer, d.ModelVersion,
				d.DeviceTypeString,
				d.SampleRate, d.Encoding, d.Latency,
				d.Redundancy, d.ExtWordClock,
				d.PriIPv4Address, d.PriIPv4Mode,
				d.SecIPv4Address, d.SecIPv4Mode,
				d.PreferredMaster, d.InteropMode, d.ClockPreferred,
			}
		}
	}

	startRow := 5
	// Write Header
	for colIdx, header := range devH {
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, startRow)
		f.SetCellValue(sheet1, cell, header)
	}
	lastHeaderCell, _ := excelize.CoordinatesToCellName(len(devH), startRow)
	f.SetCellStyle(sheet1, "A5", lastHeaderCell, headerStyleID)

	// Write Data
	rowIdx := startRow + 1
	for _, d := range deviceRows {
		vals := devValuesFunc(d)
		for colIdx, v := range vals {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx)
			f.SetCellStr(sheet1, cell, v)
		}
		rowIdx++
	}

	// Apply text and border styles to entire data table
	lastDataCell, _ := excelize.CoordinatesToCellName(len(devH), rowIdx-1)
	if rowIdx-1 >= startRow+1 {
		f.SetCellStyle(sheet1, "A6", lastDataCell, borderStyleID)
	}

	autoFitColumns(f, sheet1, devH, startRow)

	// -------------------------------------------------------------
	// SHEET 2: Patch Matrix
	// -------------------------------------------------------------
	fmt.Println("  Writing sheet: Patch Matrix...")
	sheet2 := "Patch Matrix"
	f.NewSheet(sheet2)

	oR := 3 // Origin Row (1-based index)
	oC := 3 // Origin Column (1-based index)

	if len(txRows) > 0 && len(rxRows) > 0 {
		// Define vertical orientation headers style
		verticalHeaderStyle1, _ := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Size: 8},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFD966"}, Pattern: 1},
			Alignment: &excelize.Alignment{TextRotation: 90, Horizontal: "center", Vertical: "center"},
			Border: []excelize.Border{
				{Type: "top", Color: "D0D0D0", Style: 1},
				{Type: "bottom", Color: "D0D0D0", Style: 1},
				{Type: "left", Color: "D0D0D0", Style: 1},
				{Type: "right", Color: "D0D0D0", Style: 1},
			},
		})

		verticalHeaderStyle2, _ := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Size: 7},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFF2CC"}, Pattern: 1},
			Alignment: &excelize.Alignment{TextRotation: 90, Horizontal: "center", Vertical: "center"},
			Border: []excelize.Border{
				{Type: "top", Color: "D0D0D0", Style: 1},
				{Type: "bottom", Color: "D0D0D0", Style: 1},
				{Type: "left", Color: "D0D0D0", Style: 1},
				{Type: "right", Color: "D0D0D0", Style: 1},
			},
		})

		rxHeaderStyle1, _ := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Size: 8},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"9BC2E6"}, Pattern: 1},
			Alignment: &excelize.Alignment{Vertical: "center"},
			Border: []excelize.Border{
				{Type: "top", Color: "D0D0D0", Style: 1},
				{Type: "bottom", Color: "D0D0D0", Style: 1},
				{Type: "left", Color: "D0D0D0", Style: 1},
				{Type: "right", Color: "D0D0D0", Style: 1},
			},
		})

		rxHeaderStyle2, _ := f.NewStyle(&excelize.Style{
			Font:      &excelize.Font{Bold: true, Size: 7},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"DDEBF7"}, Pattern: 1},
			Alignment: &excelize.Alignment{Vertical: "center"},
			Border: []excelize.Border{
				{Type: "top", Color: "D0D0D0", Style: 1},
				{Type: "bottom", Color: "D0D0D0", Style: 1},
				{Type: "left", Color: "D0D0D0", Style: 1},
				{Type: "right", Color: "D0D0D0", Style: 1},
			},
		})

		greenStyleID, _ := f.NewStyle(&excelize.Style{
			Fill: excelize.Fill{Type: "pattern", Color: []string{"C6EFCE"}, Pattern: 1},
			Border: []excelize.Border{
				{Type: "top", Color: "D0D0D0", Style: 1},
				{Type: "bottom", Color: "D0D0D0", Style: 1},
				{Type: "left", Color: "D0D0D0", Style: 1},
				{Type: "right", Color: "D0D0D0", Style: 1},
			},
		})

		// Apply border styles to the entire matrix workspace first
		lastColLetter, _ := excelize.ColumnNumberToName(oC + len(txRows) - 1)
		lastRowNumber := oR + len(rxRows) - 1
		f.SetCellStyle(sheet2, "A1", fmt.Sprintf("%s%d", lastColLetter, lastRowNumber), borderStyleID)

		// Column headers (TX Channels)
		curDev := ""
		startCol := oC
		maxDevLen := 0

		for t := 0; t < len(txRows); t++ {
			col := oC + t
			tx := txRows[t]

			cellLabel, _ := excelize.CoordinatesToCellName(col, 2)
			f.SetCellStr(sheet2, cellLabel, tx.Label)
			f.SetCellStyle(sheet2, cellLabel, cellLabel, verticalHeaderStyle2)

			if len(tx.Device) > maxDevLen {
				maxDevLen = len(tx.Device)
			}

			if tx.Device != curDev {
				if curDev != "" && (col-1) >= startCol {
					startCell, _ := excelize.CoordinatesToCellName(startCol, 1)
					endCell, _ := excelize.CoordinatesToCellName(col-1, 1)
					f.MergeCell(sheet2, startCell, endCell)
				}
				cellDev, _ := excelize.CoordinatesToCellName(col, 1)
				f.SetCellStr(sheet2, cellDev, tx.Device)
				f.SetCellStyle(sheet2, cellDev, cellDev, verticalHeaderStyle1)

				curDev = tx.Device
				startCol = col
			}
		}
		// Merge last group
		lastTxCol := oC + len(txRows) - 1
		if lastTxCol >= startCol {
			startCell, _ := excelize.CoordinatesToCellName(startCol, 1)
			endCell, _ := excelize.CoordinatesToCellName(lastTxCol, 1)
			f.MergeCell(sheet2, startCell, endCell)
		}

		// Adjust Row 1 height based on longest device name
		row1Height := float64(maxDevLen) * 6.5
		if row1Height < 50.0 {
			row1Height = 50.0
		}
		f.SetRowHeight(sheet2, 1, row1Height)
		f.SetRowHeight(sheet2, 2, 60.0) // Label row height

		// Row headers (RX Channels)
		curDev = ""
		startRowPatch := oR

		for i := 0; i < len(rxRows); i++ {
			row := oR + i
			rx := rxRows[i]

			cellName, _ := excelize.CoordinatesToCellName(2, row)
			f.SetCellStr(sheet2, cellName, rx.Name)
			f.SetCellStyle(sheet2, cellName, cellName, rxHeaderStyle2)

			if rx.Device != curDev {
				if curDev != "" && (row-1) >= startRowPatch {
					startCell, _ := excelize.CoordinatesToCellName(1, startRowPatch)
					endCell, _ := excelize.CoordinatesToCellName(1, row-1)
					f.MergeCell(sheet2, startCell, endCell)
				}
				cellDev, _ := excelize.CoordinatesToCellName(1, row)
				f.SetCellStr(sheet2, cellDev, rx.Device)
				f.SetCellStyle(sheet2, cellDev, cellDev, rxHeaderStyle1)

				curDev = rx.Device
				startRowPatch = row
			}
		}
		// Merge last group
		lastRxRow := oR + len(rxRows) - 1
		if lastRxRow >= startRowPatch {
			startCell, _ := excelize.CoordinatesToCellName(1, startRowPatch)
			endCell, _ := excelize.CoordinatesToCellName(1, lastRxRow)
			f.MergeCell(sheet2, startCell, endCell)
		}

		// Index map for Tx keys
		txKeyIndex := make(map[string]int)
		for t, tx := range txRows {
			txKeyIndex[tx.Key] = t
		}

		// Mark cross points
		for i, rx := range rxRows {
			if rx.SubscribedDevice != "" && rx.SubscribedChannel != "" {
				txKey := rx.SubscribedDevice + Sep + rx.SubscribedChannel
				if tIdx, ok := txKeyIndex[txKey]; ok {
					cR := oR + i
					cC := oC + tIdx
					cell, _ := excelize.CoordinatesToCellName(cC, cR)
					f.SetCellStyle(sheet2, cell, cell, greenStyleID)
				}
			}
		}

		// Final layout styling
		f.SetCellStr(sheet2, "A1", "RX / TX")
		cornerStyle, _ := f.NewStyle(&excelize.Style{
			Font: &excelize.Font{Bold: true},
			Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		})
		f.SetCellStyle(sheet2, "A1", "A1", cornerStyle)

		f.SetColWidth(sheet2, "A", "A", 18)
		f.SetColWidth(sheet2, "B", "B", 16)

		// Set column widths of matrix to 3
		for t := 0; t < len(txRows); t++ {
			colLetter, _ := excelize.ColumnNumberToName(oC + t)
			f.SetColWidth(sheet2, colLetter, colLetter, 3.5)
		}

		// Set row heights of matrix to 13
		for i := 0; i < len(rxRows); i++ {
			f.SetRowHeight(sheet2, oR+i, 14.0)
		}

		// Freeze Panes
		f.SetPanes(sheet2, &excelize.Panes{
			Freeze:      true,
			Split:       false,
			XSplit:      2,
			YSplit:      2,
			TopLeftCell: "C3",
			ActivePane:  "bottomRight",
		})

	} else {
		f.SetCellValue(sheet2, "A1", "(No TX/RX channels to display)")
	}

	// -------------------------------------------------------------
	// SHEET 3: TX Flows
	// -------------------------------------------------------------
	fmt.Println("  Writing sheet: TX Flows...")
	sheet3 := "TX Flows"
	f.NewSheet(sheet3)
	f.SetColStyle(sheet3, "A:ZZ", textStyleID)

	var flowH []string
	if detailMode {
		flowH = []string{"Device Name", "Flow Type", "Dante ID", "FPP", "Media Type", "Session ID", "Transport Type", "Dest Address", "Dest Port", "Slot Count", "Slot Channels"}
	} else {
		flowH = []string{"Device Name", "Flow Type", "Dest Address", "Dest Port", "Slot Count", "Slot Channels"}
	}

	// Write Header
	for colIdx, header := range flowH {
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
		f.SetCellValue(sheet3, cell, header)
	}
	lastFlowHeaderCell, _ := excelize.CoordinatesToCellName(len(flowH), 1)
	f.SetCellStyle(sheet3, "A1", lastFlowHeaderCell, headerStyleID)

	// Write Data
	rowIdx = 2
	for _, flow := range txFlowRows {
		if detailMode {
			f.SetCellStr(sheet3, fmt.Sprintf("A%d", rowIdx), flow.Device)
			f.SetCellStr(sheet3, fmt.Sprintf("B%d", rowIdx), flow.FlowType)
			f.SetCellStr(sheet3, fmt.Sprintf("C%d", rowIdx), flow.DanteId)
			f.SetCellStr(sheet3, fmt.Sprintf("D%d", rowIdx), flow.Fpp)
			f.SetCellStr(sheet3, fmt.Sprintf("E%d", rowIdx), flow.MediaType)
			f.SetCellStr(sheet3, fmt.Sprintf("F%d", rowIdx), flow.SessionId)
			f.SetCellStr(sheet3, fmt.Sprintf("G%d", rowIdx), flow.TransportType)
			f.SetCellStr(sheet3, fmt.Sprintf("H%d", rowIdx), flow.DestAddress)
			f.SetCellStr(sheet3, fmt.Sprintf("I%d", rowIdx), flow.DestPort)
			f.SetCellStr(sheet3, fmt.Sprintf("J%d", rowIdx), strconv.Itoa(flow.SlotCount))
			f.SetCellStr(sheet3, fmt.Sprintf("K%d", rowIdx), flow.SlotChannels)
		} else {
			f.SetCellStr(sheet3, fmt.Sprintf("A%d", rowIdx), flow.Device)
			f.SetCellStr(sheet3, fmt.Sprintf("B%d", rowIdx), flow.FlowType)
			f.SetCellStr(sheet3, fmt.Sprintf("C%d", rowIdx), flow.DestAddress)
			f.SetCellStr(sheet3, fmt.Sprintf("D%d", rowIdx), flow.DestPort)
			f.SetCellStr(sheet3, fmt.Sprintf("E%d", rowIdx), strconv.Itoa(flow.SlotCount))
			f.SetCellStr(sheet3, fmt.Sprintf("F%d", rowIdx), flow.SlotChannels)
		}
		rowIdx++
	}

	if len(txFlowRows) == 0 {
		f.SetCellStr(sheet3, "A2", "(No TX Flows)")
		rowIdx++
	}

	// Apply text & border style
	lastFlowDataCell, _ := excelize.CoordinatesToCellName(len(flowH), rowIdx-1)
	f.SetCellStyle(sheet3, "A2", lastFlowDataCell, borderStyleID)

	autoFitColumns(f, sheet3, flowH, 1)

	// -------------------------------------------------------------
	// DETAIL MODE SHEETS
	// -------------------------------------------------------------
	if detailMode {
		// SHEET 4: TX Channels
		fmt.Println("  Writing sheet: TX Channels...")
		sheet4 := "TX Channels"
		f.NewSheet(sheet4)
		f.SetColStyle(sheet4, "A:ZZ", textStyleID)

		txH := []string{"Device Name", "Dante ID", "Channel Label", "Media Type"}
		for colIdx, header := range txH {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheet4, cell, header)
		}
		lastTxHeaderCell, _ := excelize.CoordinatesToCellName(len(txH), 1)
		f.SetCellStyle(sheet4, "A1", lastTxHeaderCell, headerStyleID)

		rowIdx = 2
		for _, tx := range txRows {
			f.SetCellStr(sheet4, fmt.Sprintf("A%d", rowIdx), tx.Device)
			f.SetCellStr(sheet4, fmt.Sprintf("B%d", rowIdx), tx.DanteId)
			f.SetCellStr(sheet4, fmt.Sprintf("C%d", rowIdx), tx.Label)
			f.SetCellStr(sheet4, fmt.Sprintf("D%d", rowIdx), tx.MediaType)
			rowIdx++
		}
		if len(txRows) == 0 {
			f.SetCellStr(sheet4, "A2", "(No TX Channels)")
			rowIdx++
		}
		lastTxDataCell, _ := excelize.CoordinatesToCellName(len(txH), rowIdx-1)
		f.SetCellStyle(sheet4, "A2", lastTxDataCell, borderStyleID)
		autoFitColumns(f, sheet4, txH, 1)

		// SHEET 5: RX Channels
		fmt.Println("  Writing sheet: RX Channels...")
		sheet5 := "RX Channels"
		f.NewSheet(sheet5)
		f.SetColStyle(sheet5, "A:ZZ", textStyleID)

		rxH := []string{"Device Name", "Dante ID", "Channel Name", "Media Type", "Subscribed Device", "Subscribed Channel", "Status"}
		for colIdx, header := range rxH {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheet5, cell, header)
		}
		lastRxHeaderCell, _ := excelize.CoordinatesToCellName(len(rxH), 1)
		f.SetCellStyle(sheet5, "A1", lastRxHeaderCell, headerStyleID)

		rowIdx = 2
		for _, rx := range rxRows {
			f.SetCellStr(sheet5, fmt.Sprintf("A%d", rowIdx), rx.Device)
			f.SetCellStr(sheet5, fmt.Sprintf("B%d", rowIdx), rx.DanteId)
			f.SetCellStr(sheet5, fmt.Sprintf("C%d", rowIdx), rx.Name)
			f.SetCellStr(sheet5, fmt.Sprintf("D%d", rowIdx), rx.MediaType)

			status := "Unsubscribed"
			if rx.SubscribedDevice != "" && rx.SubscribedChannel != "" {
				f.SetCellStr(sheet5, fmt.Sprintf("E%d", rowIdx), rx.SubscribedDevice)
				f.SetCellStr(sheet5, fmt.Sprintf("F%d", rowIdx), rx.SubscribedChannel)
				status = "Connected"
			}
			f.SetCellStr(sheet5, fmt.Sprintf("G%d", rowIdx), status)
			rowIdx++
		}
		if len(rxRows) == 0 {
			f.SetCellStr(sheet5, "A2", "(No RX Channels)")
			rowIdx++
		}
		lastRxDataCell, _ := excelize.CoordinatesToCellName(len(rxH), rowIdx-1)
		f.SetCellStyle(sheet5, "A2", lastRxDataCell, borderStyleID)
		autoFitColumns(f, sheet5, rxH, 1)

		// SHEET 6: Subscriptions
		fmt.Println("  Writing sheet: Subscriptions...")
		sheet6 := "Subscriptions"
		f.NewSheet(sheet6)
		f.SetColStyle(sheet6, "A:ZZ", textStyleID)

		subH := []string{"No.", "RX Device", "RX Channel", "RX Dante ID", "TX Device", "TX Channel", "Media Type"}
		for colIdx, header := range subH {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheet6, cell, header)
		}
		lastSubHeaderCell, _ := excelize.CoordinatesToCellName(len(subH), 1)
		f.SetCellStyle(sheet6, "A1", lastSubHeaderCell, headerStyleID)

		rowIdx = 2
		for no, sub := range subRows {
			f.SetCellStr(sheet6, fmt.Sprintf("A%d", rowIdx), strconv.Itoa(no+1))
			f.SetCellStr(sheet6, fmt.Sprintf("B%d", rowIdx), sub.RxDevice)
			f.SetCellStr(sheet6, fmt.Sprintf("C%d", rowIdx), sub.RxChannel)
			f.SetCellStr(sheet6, fmt.Sprintf("D%d", rowIdx), sub.RxDanteId)
			f.SetCellStr(sheet6, fmt.Sprintf("E%d", rowIdx), sub.TxDevice)
			f.SetCellStr(sheet6, fmt.Sprintf("F%d", rowIdx), sub.TxChannel)
			f.SetCellStr(sheet6, fmt.Sprintf("G%d", rowIdx), sub.MediaType)
			rowIdx++
		}
		if len(subRows) == 0 {
			f.SetCellStr(sheet6, "A2", "(No Subscriptions)")
			rowIdx++
		}
		lastSubDataCell, _ := excelize.CoordinatesToCellName(len(subH), rowIdx-1)
		f.SetCellStyle(sheet6, "A2", lastSubDataCell, borderStyleID)
		autoFitColumns(f, sheet6, subH, 1)
	}

	// Make sheet "Devices" active
	f.SetActiveSheet(0)

	// Save
	fmt.Println("Saving...")
	if err := f.SaveAs(outputXlsx); err != nil {
		fmt.Printf("[ERROR] Failed to save Excel: %v\n", err)
		waitForEnter()
		os.Exit(1)
	}

	modeLabel := "Default"
	if detailMode {
		modeLabel = "Detail"
	}

	fmt.Println("")
	fmt.Println("====================================")
	fmt.Printf(" Done: %s\n", outputXlsx)
	fmt.Printf(" Mode:          %s\n", modeLabel)
	fmt.Printf(" Devices:       %d\n", len(preset.Devices))
	fmt.Printf(" TX Channels:   %d\n", len(txRows))
	fmt.Printf(" TX Flows:      %d\n", len(txFlowRows))
	fmt.Printf(" RX Channels:   %d\n", len(rxRows))
	fmt.Printf(" Subscriptions: %d\n", len(subRows))
	fmt.Println("====================================")

	waitForEnter()
}

func selectXMLFile() (string, error) {
	files, err := filepath.Glob("*.xml")
	if err != nil {
		return "", err
	}

	if len(files) == 0 {
		return "", fmt.Errorf("no .xml files found in the current directory")
	}

	if len(files) == 1 {
		return files[0], nil
	}

	fmt.Println("=== XML files found ===")
	for i, f := range files {
		fmt.Printf("  %d: %s\n", i+1, f)
	}
	fmt.Println("")

	scanner := bufio.NewScanner(os.Stdin)
	for {
		fmt.Printf("Select number (1-%d): ", len(files))
		if !scanner.Scan() {
			return "", fmt.Errorf("input cancelled")
		}
		choice := strings.TrimSpace(scanner.Text())
		idx, err := strconv.Atoi(choice)
		if err != nil || idx < 1 || idx > len(files) {
			fmt.Println("Invalid selection. Please try again.")
			continue
		}
		return files[idx-1], nil
	}
}

func selectMode() (bool, error) {
	scanner := bufio.NewScanner(os.Stdin)
	for {
		fmt.Println("=== Menu ===")
		fmt.Println("  1: Default (summary)")
		fmt.Println("  2: Detail  (all info)")
		fmt.Println("  H: Help")
		fmt.Println("")
		fmt.Print("Select (1 / 2 / H) [default: 1]: ")

		if !scanner.Scan() {
			return false, fmt.Errorf("input cancelled")
		}
		choice := strings.ToLower(strings.TrimSpace(scanner.Text()))

		if choice == "" || choice == "1" {
			fmt.Println("-> Default mode")
			return false, nil
		} else if choice == "2" {
			fmt.Println("-> Detail mode")
			return true, nil
		} else if choice == "h" {
			fmt.Println("")
			fmt.Println(helpText)
			fmt.Println("")
			continue
		} else {
			fmt.Println("Invalid selection. Please try again.")
		}
	}
}

func autoFitColumns(f *excelize.File, sheetName string, headers []string, startRow int) {
	colWidths := make(map[int]int)

	// Initialize width from headers
	for colIdx, h := range headers {
		colWidths[colIdx+1] = len(h)
	}

	rows, err := f.GetRows(sheetName)
	if err == nil {
		for rIdx := startRow - 1; rIdx < len(rows); rIdx++ {
			for cIdx, val := range rows[rIdx] {
				l := len(val)
				if l > colWidths[cIdx+1] {
					colWidths[cIdx+1] = l
				}
			}
		}
	}

	// Apply widths
	for col, width := range colWidths {
		colName, err := excelize.ColumnNumberToName(col)
		if err == nil {
			w := float64(width) + 3.0
			if w > 50.0 {
				w = 50.0
			}
			if w < 10.0 {
				w = 10.0
			}
			f.SetColWidth(sheetName, colName, colName, w)
		}
	}
}

func waitForEnter() {
	fmt.Println("")
	fmt.Println("Press Enter to exit...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
}
