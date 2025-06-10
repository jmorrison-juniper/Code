📡 MistHelper
MistHelper is a comprehensive Python-based CLI and automation toolkit designed for Juniper Mist environments. It enables NOC engineers to efficiently extract, analyze, and interact with network data across sites, devices, and configurations using Mist APIs.

🚀 Features

🔍 Audit & Alarm Logs: Export alarms, audit logs, and event definitions.

🧠 Marvis Actions: Poll and export open Marvis AI actions.

🧰 Device & Site Inventory: Export full org inventory, site lists, and enriched device metadata.

📊 Statistics & Port Data: Gather device stats, port-level metrics, and VPN peer stats.

🧪 Synthetic Tests: Collect synthetic test results from gateways.

🧵 Virtual Chassis: Export switch stacking and VC stats.

🖥️ Interactive CLI Shell: Launch WebSocket-based shell sessions to run commands like show route, show vlans, etc.

🧾 Support Package Generator: Automatically compile per-site support packages with alarms, events, and performance data.

🧠 Dynamic API Rate Control: PID-based delay tuning to avoid hitting Mist API rate limits.

🧱 Requirements

The script auto-installs required packages if missing:


🔧 Setup

Clone the repo or copy the script to your working directory.

Create a .env file with your Mist credentials:

Run the script:

🖥️ Usage Modes

🔹 Interactive Mode

Run without arguments to access a menu-driven interface:


🔹 CLI Mode

Run specific tasks directly:


📋 Menu Options Overview

Option	Description

1	Export all org alarms (last 24h)

2	Export all device events

11	Export site list

12	Export org device inventory

16	View site device inventory

20	Export synthetic test results for all gateways

29	Generate support packages per site

33	Launch interactive CLI shell

34	Run ARP command via WebSocket

38	Run show route 0.0.0.0 and export to CSV

40	Run show vlans and export to CSV

📦 Support Package

Generates per-site CSVs containing:

Alarms

Device events

Device stats

Port stats

Speed test results

Useful for escalations, audits, or proactive monitoring.

🧠 Smart API Rate Handling

The script uses a PID controller to dynamically adjust API call delays based on usage trends, ensuring compliance with Mist API rate limits.

🛠️ Advanced Capabilities

WebSocket Shell: Real-time CLI interaction with Mist devices.

ARP Command Streaming: Trigger and stream ARP table output via WebSocket.

Dynamic CSV Merging: Combine SFP module data with site/device metadata.

JSON-to-CSV Conversion: Extract and flatten JSON from CLI output logs.

📁 Output Files

The script generates multiple CSVs including:

OrgAlarms.csv

OrgDeviceStats.csv

AllDevicesWithSiteInfo.csv

AllGatewaySyntheticTests.csv

SupportPackage_<site_id>.csv

🧪 Example: Run a Shell Command

This runs show route 0.0.0.0 on the selected device and saves the output to RouteDefault.csv.

🧼 Cleanup

To stop continuous loops or refreshes, create a file named stop_loop.txt in the working directory.

📞 Support

For questions or issues, contact your network automation team or Mist API administrator.