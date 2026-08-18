# 🌐 Network Engineer Toolkit

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![UI: customtkinter](https://img.shields.io/badge/UI-customtkinter-blueviolet)](https://github.com/TomSchimansky/CustomTkinter)

An extensible Python suite for network automation and engineering. Built for reliability, speed, and enterprise workflows.

## 🧰 Tools Included

| Tool | Version | Status | Description |
|------|---------|--------|-------------|
| **IP Subnet Calculator** | v1.0 | ✅ Active | High-performance IPv4/IPv6 subnetting, VLSM calculation, and smart next-hop logic with a modern Dark Mode GUI. |
| **Bulk Network Ping Monitor** | v1.0 | ✅ Active | High-speed, parallel ping monitoring for hundreds of devices with a live dark-mode dashboard, Excel integration, and automated CSV reporting. |
| **Config Backup Runner** | v1.0 | ✅ Active | Pings, then SSHes into every device in `devices.xlsx` (Cisco IOS and Huawei VRP), saves each running-config to a dated `backups/` folder, and writes an OK/FAILED CSV report. Read-only; credentials prompted at run time, never stored. |
... In progress ...

## 🚀 Installation & Setup

1. Clone the repository to your local machine:
```bash
git clone [https://github.com/jijoeee/network-engineer-toolkit.git](https://github.com/jijoeee/network-engineer-toolkit.git)
cd network-engineer-toolkit
```
2. Install the required GUI dependencies:

```Bash

pip install -r requirements.txt
```
3. Navigate to the tool's directory and launch the GUI:

```Bash

cd tools/subnet_calculator
python ip_subnet_calculator.py
```

## ✨ Current Features

1) IP Subnet Calculator

    IPv4 & IPv6 Support: Handles both protocols seamlessly using Python's native ipaddress library.
    
    Smart Next-Hop Logic: Accurately calculates next-hops, specifically accounting for /31 and /127 point-to-point links.
    
    Subnet Splitting Engine: Slice large blocks (e.g., /24) into smaller prefixes (e.g., /28) using a low-memory generator algorithm.
    
    Modern Interface: Built with customtkinter for a professional, dark-mode-first user experience.

2) Bulk Network Ping Monitor
    
    Parallel Execution Engine: Uses Python's ThreadPoolExecutor to simultaneously ping hundreds of devices, dropping scan times from minutes to seconds.
    
    Smart Excel Integration: Directly imports device inventory (Hostname, IP, Location, Type, Group) from devices.xlsx without rigid header requirements.
    
    Advanced Analytics Tracking: Continuously tracks and calculates Packet Loss %, Uptime, Last Down Time, and Total Ping Counts per session.
    
    Dynamic Smart Filtering: Instantly filter the live dashboard by Device Group, Location, Type, or UP/DOWN status.
    
    Session Management: Features a clean Pause/Resume engine and automatically resets data counters upon starting a fresh session.
    
    Automated NOC Reporting: One-click export dynamically generates a timestamped CSV report and saves it directly to a dedicated report/ directory.

3) Config Backup Runner

    Same Inventory: Reads the very same devices.xlsx (Hostname, IP, Location, Device Type, Group) the Ping Monitor uses — one device list for the whole toolkit.

    Ping-First, Then SSH: Pings each device before connecting, so unreachable boxes are skipped instead of blocking on a 30-second SSH timeout.

    Multi-Vendor: Pulls the running-config from Cisco IOS/IOS-XE (show running-config) and Huawei VRP (display current-configuration); Arista, HP Comware and Juniper are recognised too, with a safe default for anything else.

    Dated Archive: Saves each config to backups/YYYY-MM-DD/<hostname>.cfg — an instant point-in-time record of every box.

    OK/FAILED Report: Writes a timestamped CSV to report/ saying exactly which devices backed up and why any failed (unreachable, auth, timeout).

    Safe by Design: Read-only — it never pushes config. Credentials are prompted at run time or read from NET_USER / NET_PASS / NET_ENABLE, and are never written to a file.
