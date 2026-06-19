package main

import (
	"encoding/xml"
	"io"
	"os"
	"testing"
)

func TestParseXML(t *testing.T) {
	// Open test XML file
	xmlFile, err := os.Open("aaaa.xml")
	if err != nil {
		t.Fatalf("Failed to open test file: %v", err)
	}
	defer xmlFile.Close()

	byteValue, err := io.ReadAll(xmlFile)
	if err != nil {
		t.Fatalf("Failed to read test file: %v", err)
	}

	var preset Preset
	err = xml.Unmarshal(byteValue, &preset)
	if err != nil {
		t.Fatalf("Failed to unmarshal XML: %v", err)
	}

	// Assertions
	if preset.Name != "aaaa" {
		t.Errorf("Expected preset name 'aaaa', got '%s'", preset.Name)
	}

	if preset.GetVersion() != "3.0.0" {
		t.Errorf("Expected preset version '3.0.0', got '%s'", preset.GetVersion())
	}

	if len(preset.Devices) != 2 {
		t.Fatalf("Expected 2 devices, got %d", len(preset.Devices))
	}

	dev0 := preset.Devices[0]
	if dev0.Name != "AVIOUSB-507d4b" {
		t.Errorf("Expected device 0 name 'AVIOUSB-507d4b', got '%s'", dev0.Name)
	}
	if dev0.SampleRate != "48000" {
		t.Errorf("Expected device 0 samplerate '48000', got '%s'", dev0.SampleRate)
	}
	if dev0.Encoding != "24" {
		t.Errorf("Expected device 0 encoding '24', got '%s'", dev0.Encoding)
	}
	if dev0.PreferredMaster.Value != "false" {
		t.Errorf("Expected device 0 preferred_master value 'false', got '%s'", dev0.PreferredMaster.Value)
	}

	if len(dev0.Interfaces) != 1 {
		t.Fatalf("Expected 1 interface for device 0, got %d", len(dev0.Interfaces))
	}
	if dev0.Interfaces[0].IPv4Address.Mode != "dynamic" {
		t.Errorf("Expected interface 0 mode 'dynamic', got '%s'", dev0.Interfaces[0].IPv4Address.Mode)
	}

	dev1 := preset.Devices[1]
	if dev1.Name != "Core-a088" {
		t.Errorf("Expected device 1 name 'Core-a088', got '%s'", dev1.Name)
	}
	if dev1.UnicastLatency != "2000" {
		t.Errorf("Expected device 1 unicast_latency '2000', got '%s'", dev1.UnicastLatency)
	}

	// Verify rxchannel subscription
	var foundSub bool
	for _, rx := range dev1.RxChannels {
		if rx.DanteId == "1" {
			if rx.Name != "01 Software-Dante-RX-1" {
				t.Errorf("Expected rxchannel 1 name '01 Software-Dante-RX-1', got '%s'", rx.Name)
			}
			if rx.SubscribedDevice != "AVIOUSB-507d4b" {
				t.Errorf("Expected subscribed device 'AVIOUSB-507d4b', got '%s'", rx.SubscribedDevice)
			}
			if rx.SubscribedChannel != "Left" {
				t.Errorf("Expected subscribed channel 'Left', got '%s'", rx.SubscribedChannel)
			}
			foundSub = true
		}
	}
	if !foundSub {
		t.Errorf("rxchannel with danteId 1 not found or verified in Core-a088")
	}
}

func TestProcessPresetData(t *testing.T) {
	xmlFile, err := os.Open("aaaa.xml")
	if err != nil {
		t.Fatalf("Failed to open test file: %v", err)
	}
	defer xmlFile.Close()

	byteValue, err := io.ReadAll(xmlFile)
	if err != nil {
		t.Fatalf("Failed to read test file: %v", err)
	}

	var preset Preset
	err = xml.Unmarshal(byteValue, &preset)
	if err != nil {
		t.Fatalf("Failed to unmarshal XML: %v", err)
	}

	deviceRows, txRows, txFlowRows, rxRows, subRows := ProcessPresetData(&preset)

	// Assertions on structured rows
	if len(deviceRows) != 2 {
		t.Errorf("Expected 2 device rows, got %d", len(deviceRows))
	}
	if deviceRows[0].DeviceName != "AVIOUSB-507d4b" {
		t.Errorf("Expected first device 'AVIOUSB-507d4b', got '%s'", deviceRows[0].DeviceName)
	}
	if deviceRows[0].SampleRate != "48000" {
		t.Errorf("Expected SampleRate '48000', got '%s'", deviceRows[0].SampleRate)
	}
	if deviceRows[0].Encoding != "24" {
		t.Errorf("Expected Encoding '24', got '%s'", deviceRows[0].Encoding)
	}

	// TX Rows: AVIOUSB-507d4b has 2 TX channels (Left, Right)
	if len(txRows) != 2 {
		t.Errorf("Expected 2 TX rows, got %d", len(txRows))
	}
	if txRows[0].Label != "Left" || txRows[0].Device != "AVIOUSB-507d4b" {
		t.Errorf("Expected first TX row: AVIOUSB-507d4b, Left. Got %s, %s", txRows[0].Device, txRows[0].Label)
	}

	// RX Rows: AVIOUSB-507d4b has 2 RX channels, Core-a088 has 8 RX channels = 10 RX rows
	if len(rxRows) != 10 {
		t.Errorf("Expected 10 RX rows, got %d", len(rxRows))
	}

	// Subscriptions: Core-a088 has 1 subscription to AVIOUSB-507d4b Left
	if len(subRows) != 1 {
		t.Errorf("Expected 1 subscription, got %d", len(subRows))
	}
	if subRows[0].RxDevice != "Core-a088" || subRows[0].TxDevice != "AVIOUSB-507d4b" || subRows[0].TxChannel != "Left" {
		t.Errorf("Unexpected subscription: Rx=%s, Tx=%s, Ch=%s", subRows[0].RxDevice, subRows[0].TxDevice, subRows[0].TxChannel)
	}

	// TX Flows: aaaa.xml has no txflow in first device or second device.
	if len(txFlowRows) != 0 {
		t.Errorf("Expected 0 TX flows, got %d", len(txFlowRows))
	}
}

func TestProcessPresetDataWithFlows(t *testing.T) {
	// Let's test with xxxx.xml which has a txflow
	xmlFile, err := os.Open("xxxx.xml")
	if err != nil {
		t.Fatalf("Failed to open test file: %v", err)
	}
	defer xmlFile.Close()

	byteValue, err := io.ReadAll(xmlFile)
	if err != nil {
		t.Fatalf("Failed to read test file: %v", err)
	}

	var preset Preset
	err = xml.Unmarshal(byteValue, &preset)
	if err != nil {
		t.Fatalf("Failed to unmarshal XML: %v", err)
	}

	_, _, txFlowRows, _, _ := ProcessPresetData(&preset)

	// xxxx.xml has 1 txflow in AVIOUSB-507d4b:
	// <txflow danteId="2" fpp="48" mediaType="audio" sessionId="27455922" transportType="2">
	//     <slot channelId="1"/>
	//     <slot channelId="2"/>
	//     <destinationAddress address="239.69.99.168" port="5004"/>
	// </txflow>
	if len(txFlowRows) != 1 {
		t.Fatalf("Expected 1 TX flow row, got %d", len(txFlowRows))
	}

	flow := txFlowRows[0]
	if flow.Device != "AVIOUSB-507d4b" {
		t.Errorf("Expected flow device 'AVIOUSB-507d4b', got '%s'", flow.Device)
	}
	if flow.FlowType != "AES67" {
		t.Errorf("Expected AES67 flow, got '%s'", flow.FlowType)
	}
	if flow.DestAddress != "239.69.99.168" || flow.DestPort != "5004" {
		t.Errorf("Expected dest 239.69.99.168:5004, got %s:%s", flow.DestAddress, flow.DestPort)
	}
	if flow.SlotCount != 2 || flow.SlotChannels != "1, 2" {
		t.Errorf("Expected 2 slots '1, 2', got %d '%s'", flow.SlotCount, flow.SlotChannels)
	}
}
