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
| **Config Backup Runner** | v1.0 | 🧪 Beta — not yet run against live hardware | Pings, then SSHes into every device in `devices.xlsx` (Cisco IOS and Huawei VRP), saves each running-config to a dated `backups/` folder, and writes an OK/FAILED CSV report. Read-only; credentials prompted at run time, never stored. |
| **Bulk Show-Command Runner** | v1.0 | 🧪 Beta — not yet run against live hardware | Runs the same list of `show` commands across every device in `devices.xlsx` in parallel and collects the answers in one Excel workbook, one sheet per command. Run it again after a change and it marks every device whose output differs from last time **and shows you the changed lines** — a pre/post-change check instead of an evening of pasted screenshots. Read-only; credentials prompted at run time, never stored. |
... In progress ...

> **On the Beta label:** the Config Backup Runner and the Bulk Show-Command Runner have their
> parsing, vendor mapping, output paths, reporting, diff logic and error handling covered by
> offline tests, but no maintainer has yet run either against a physical Cisco or Huawei device.
> Try them on one lab box before you point them at a production inventory. Both are read-only by
> design — neither has a code path that writes to a device.

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

    Same Inventory: Reads the very same devices.xlsx (Hostname, IP, Location, Device Type, Group) the Ping Monitor uses — one device list for the whole toolkit. A ready-made sample ships in the tool's folder, so it runs out of the box.

    Ping-First, Then SSH: Pings each device before connecting, so unreachable boxes are skipped instead of blocking on a 30-second SSH timeout.

    Multi-Vendor: Pulls the running-config from Cisco IOS/IOS-XE (show running-config) and Huawei VRP (display current-configuration); Arista, HP Comware and Juniper are recognised too. The vendor is read from the Device Type column and is matched inside the cell, so "Huawei S5720" and "Cisco IOS-XE router" both resolve — you do not have to reformat your existing inventory.

    Tells You Before, Not After: A bare "router" or "switch" in the Device Type column names the shape of the box, not the CLI it speaks. Those fall back to Cisco, and every device that fell back is listed on screen before the run starts — so you fix one spreadsheet column instead of reading 200 FAILED rows.

    Dated Archive: Saves each config to backups/YYYY-MM-DD/<hostname>.cfg — an instant point-in-time record of every box.

    OK/FAILED Report: Writes a timestamped CSV to report/ saying exactly which devices backed up and why any failed (unreachable, auth, timeout).

    Safe by Design: Read-only — it never pushes config. Credentials are prompted at run time or read from NET_USER / NET_PASS / NET_ENABLE, are never written to a file, and are masked out of any error message before it reaches the log or the CSV.

### How to use the Config Backup Runner

```bash
cd tools/config_backup_runner

# 1. Open devices.xlsx and put your own devices in it. Five columns, header on row 1:
#
#      Hostname  | IP Address | Location | Device Type (vendor) | Group
#      core-rtr1 | 10.0.0.1   | HQ       | cisco_ios            | core
#      dist-sw1  | 10.0.0.2   | HQ       | huawei               | distribution
#
#    Device Type must name the VENDOR, not the form factor: cisco / cisco_xe /
#    huawei / arista / juniper / hp_comware. Anything the tool cannot recognise is
#    listed as a warning before the run starts.

# 2. Run it. It asks for the SSH username and password, and they are never saved:
python config_backup_runner.py

#    To run it unattended instead (a scheduled nightly backup), export them first:
#      export NET_USER=netadmin NET_PASS='...' NET_ENABLE='...'    # Linux/macOS
#      set NET_USER=netadmin & set NET_PASS=...                    # Windows cmd

# 3. Read the results:
#      backups/2026-08-19/core-rtr1.cfg          one file per device that succeeded
#      report/Config_Backup_Report_*.csv         OK/FAILED and the reason, per device
```

Both `backups/` and `report/` are git-ignored — your configs are yours and never get committed.

4) Bulk Show-Command Runner

    Same Inventory Again: Reads the same devices.xlsx as the Ping Monitor and the Config Backup Runner. One device list drives the whole toolkit, and a ready-made sample ships in the tool's folder.

    Your Commands, Not Ours: The commands to run live in a plain commands.txt, one per line, with # for comments. A shipped sample carries safe read-only Cisco IOS checks plus commented-out Huawei VRP equivalents. Nothing is hard-coded, so the tool is as useful for a Huawei access layer as for a Cisco core.

    Parallel, Ping-First: Pings each device before connecting and skips the dead ones, then runs every command on every reachable device across a pool of 10 SSH sessions — modest on purpose, because a TACACS/RADIUS back-end dislikes a stampede.

    One Workbook, One Tab Per Command: Writes report/Show_Report_YYYYMMDD_HHMM.xlsx with a Summary sheet (per device: OK/FAILED and why) and one sheet per command holding every device's output side by side.

    Tells You What Changed, Not Just That It Did: Every run saves a snapshot. The next run compares against it and each command sheet gains a "Changed?" column (CHANGED / SAME / NEW) and a "What changed" column containing the actual added and removed lines. Changed rows are highlighted amber. That is the difference between this and a for-loop: after a change window you open one workbook and read the diff instead of eyeballing 200 outputs.

    Sensible About Noise: Trailing whitespace alone never counts as a change; an over-chatty device is capped so one box cannot produce an unreadable cell; and if a whole run fails (wrong password, management VLAN down) no snapshot is written at all, so your last good baseline survives instead of being replaced by a file full of blanks.

    Safe by Design: Read-only — the only thing it ever sends a device is a command from your commands.txt, and it has no config-write code path. Credentials are prompted at run time or read from NET_USER / NET_PASS / NET_ENABLE, are never written to a file, and are masked out of any error message before it reaches the screen or the workbook.

### How to use the Bulk Show-Command Runner

A worked example — a pre/post-change check across the whole fleet.

```bash
cd tools/bulk_show_runner

# 1. devices.xlsx — the same five columns as every other tool, header on row 1.
#    A sample ships here; replace the rows with your own:
#
#      hostname  | ip address    | location  | device type (vendor) | group
#      core-rtr1 | 192.0.2.1     | KL-DC     | cisco_ios            | 1
#      access-sw | 192.0.2.2     | KL-Office | huawei               | 1
#      edge-rtr2 | 198.51.100.10 | Cloud     | cisco_ios            | 3
#
#    "device type" must name the VENDOR (cisco / cisco_xe / huawei / arista /
#    juniper / hp_comware), not the form factor. A bare "router" or "switch" is
#    listed as a warning BEFORE the run starts, so you fix one spreadsheet
#    column instead of reading 200 FAILED rows.

# 2. commands.txt — one command per line, # to comment out. Shipped as:
#
#      # --- Cisco IOS examples ---
#      show ip interface brief
#      show cdp neighbors
#      show version
#      show ip route summary
#      # show interfaces status
#
#    Only put read-only show/display commands in here.

# 3. BEFORE your change window, take the baseline:
python bulk_show_runner.py

#    It asks for the SSH username and password (never saved), pings, connects,
#    and prints a line per device:
#
#      Loaded 3 device(s) from devices.xlsx.
#      Running 4 command(s) per device:
#          - show ip interface brief
#          ...
#      SSH username: netadmin
#      SSH password:
#      ------------------------------------------------------------
#      [OK  ] core-rtr1            192.0.2.1        4 command(s) run
#      [OK  ] access-sw            192.0.2.2        4 command(s) run
#      [FAIL] edge-rtr2            198.51.100.10    unreachable (ping failed) — skipped
#      ------------------------------------------------------------
#      Done: 2 OK, 1 failed.
#      Workbook -> report/Show_Report_20260820_0210.xlsx
#      Snapshot -> snapshots/snapshot_20260820_021044.json

# 4. Make your change. Then run exactly the same command again:
python bulk_show_runner.py

#    This time it says which baseline it is comparing against, and the workbook
#    gains the two diff columns:
#
#      Comparing against previous run: snapshot_20260820_021044.json
#      ...
#      Rows highlighted amber in the workbook changed since the last run;
#      the 'What changed' column says how.

# 5. Open report/Show_Report_20260820_0355.xlsx:
#
#      Summary tab            core-rtr1  OK  4 command(s) run   Commands changed: 1
#      "show ip interface brief" tab, core-rtr1 row, amber:
#          Changed?      CHANGED
#          What changed  -1 / +2 line(s)
#                        -GigabitEthernet0/2   unassigned   YES unset  down    down
#                        +GigabitEthernet0/2   10.20.30.1   YES manual up      up
#                        +GigabitEthernet0/3   unassigned   YES unset  down    down
#          Output        (the full command output for that device)
#
#    SAME = identical to last run. NEW = device or command not in the baseline.

# To run it unattended (a nightly fleet snapshot), export the credentials first:
#      export NET_USER=netadmin NET_PASS='...' NET_ENABLE='...'    # Linux/macOS
#      set NET_USER=netadmin & set NET_PASS=...                    # Windows cmd

# Options:
python bulk_show_runner.py -c precheck.txt   # a different command list
python bulk_show_runner.py --no-compare      # just collect, skip the diff
```

`report/` and `snapshots/` are git-ignored — your device output is yours and never gets committed.
