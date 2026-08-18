#!/usr/bin/env python3
"""
Config Backup Runner — Network Engineer Toolkit
================================================

Backs up the running-config of every device in your inventory into a dated
folder, and tells you which ones failed. Reuses the same `devices.xlsx` layout
as the Bulk Ping Monitor (Hostname, IP Address, Location, Device Type, Group).

What it does, in order:
  1. Reads the device list from `devices.xlsx`.
  2. Pings each device first, so dead boxes are skipped instead of blocking on a
     30-second SSH timeout.
  3. Logs in over SSH (Netmiko), pulls the running-config
       - Cisco IOS / IOS-XE : `show running-config`
       - Huawei VRP         : `display current-configuration`
     and writes it to `backups/YYYY-MM-DD/<hostname>.cfg`.
  4. Writes a timestamped CSV report to `report/` that says OK or FAILED,
     with the reason, for every device.

Credentials are asked for at run time, or read from the environment
(NET_USER / NET_PASS / NET_ENABLE). They are never written to a file.

Vendor is taken from the "Device Type" column in the spreadsheet. Recognised
values (case-insensitive): cisco, cisco_ios, ios, huawei, hp_comware, arista,
juniper. Anything unrecognised defaults to cisco_ios and is flagged in the log.

Usage:
    pip install -r requirements.txt      # netmiko, openpyxl
    cd tools/config_backup_runner
    python config_backup_runner.py       # uses ./devices.xlsx

Nothing here pushes config to a device — it is read-only by design.
"""

import os
import sys
import csv
import time
import getpass
import platform
import subprocess
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Dependencies -----------------------------------------------------------
try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl is required -> pip install openpyxl")
    sys.exit(1)

try:
    from netmiko import ConnectHandler
    from netmiko.exceptions import (
        NetmikoAuthenticationException,
        NetmikoTimeoutException,
    )
except ImportError:
    print("ERROR: netmiko is required -> pip install netmiko")
    sys.exit(1)


# --- Path resolution (matches the rest of the toolkit) ----------------------
def get_base_dir():
    """Folder the script/exe lives in, so it works both as .py and frozen .exe."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_dir()
EXCEL_FILE = os.path.join(BASE_DIR, "devices.xlsx")

# How many devices to back up at once. Kept modest — SSH sessions are heavier
# than pings, and TACACS/RADIUS back-ends dislike a stampede.
MAX_WORKERS = 10
PING_TIMEOUT_S = 2

# Map the human "Device Type" cell to a Netmiko device_type and the command
# that dumps the running-config for that platform.
VENDOR_MAP = {
    "cisco":       ("cisco_ios",   "show running-config"),
    "cisco_ios":   ("cisco_ios",   "show running-config"),
    "ios":         ("cisco_ios",   "show running-config"),
    "cisco_xe":    ("cisco_xe",    "show running-config"),
    "huawei":      ("huawei",      "display current-configuration"),
    "hp_comware":  ("hp_comware",  "display current-configuration"),
    "arista":      ("arista_eos",  "show running-config"),
    "arista_eos":  ("arista_eos",  "show running-config"),
    "juniper":     ("juniper_junos", "show configuration | display set"),
}
DEFAULT_VENDOR = ("cisco_ios", "show running-config")


def resolve_vendor(device_type):
    """Return (netmiko_device_type, backup_command, was_defaulted)."""
    key = (device_type or "").strip().lower().replace(" ", "_")
    if key in VENDOR_MAP:
        dt, cmd = VENDOR_MAP[key]
        return dt, cmd, False
    return DEFAULT_VENDOR[0], DEFAULT_VENDOR[1], True


# --- Inventory --------------------------------------------------------------
def load_devices():
    """Read devices.xlsx exactly like the Bulk Ping Monitor does (row 1 = header)."""
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: could not find '{EXCEL_FILE}'")
        print("Put a devices.xlsx next to this script with columns: "
              "Hostname, IP Address, Location, Device Type, Group")
        sys.exit(1)

    devices = []
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        if not row or not row[0] or not row[1]:
            continue
        devices.append({
            "hostname":    str(row[0]).strip(),
            "ip":          str(row[1]).strip(),
            "location":    str(row[2]).strip() if len(row) > 2 and row[2] else "-",
            "device_type": str(row[3]).strip() if len(row) > 3 and row[3] else "-",
            "group":       str(row[4]).strip() if len(row) > 4 and row[4] else "-",
        })
    return devices


# --- Reachability -----------------------------------------------------------
def is_reachable(ip):
    """One cheap ping so we skip dead devices instead of waiting on SSH timeouts."""
    if platform.system().lower() == "windows":
        cmd = ["ping", "-n", "1", "-w", str(PING_TIMEOUT_S * 1000), ip]
    else:
        cmd = ["ping", "-c", "1", "-W", str(PING_TIMEOUT_S), ip]
    try:
        completed = subprocess.run(
            cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        return completed.returncode == 0
    except Exception:
        return False


# --- Backup one device ------------------------------------------------------
def backup_device(device, creds, out_dir):
    """Return a result dict for the CSV report."""
    hostname = device["hostname"]
    ip = device["ip"]
    result = {
        "hostname": hostname, "ip": ip, "device_type": device["device_type"],
        "status": "FAILED", "detail": "", "file": "",
    }

    if not is_reachable(ip):
        result["detail"] = "unreachable (ping failed) — skipped"
        return result

    dt, cmd, defaulted = resolve_vendor(device["device_type"])
    if defaulted:
        result["detail"] = f"unknown device type '{device['device_type']}', assumed cisco_ios; "

    conn_params = {
        "device_type": dt,
        "host": ip,
        "username": creds["username"],
        "password": creds["password"],
        "conn_timeout": 15,
        "fast_cli": False,
    }
    if creds.get("secret"):
        conn_params["secret"] = creds["secret"]

    try:
        with ConnectHandler(**conn_params) as conn:
            if creds.get("secret"):
                try:
                    conn.enable()
                except Exception:
                    pass  # not all platforms/accounts need enable
            output = conn.send_command(cmd, read_timeout=90)

        if not output or not output.strip():
            result["detail"] += "connected but config was empty"
            return result

        out_path = os.path.join(out_dir, f"{hostname}.cfg")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(output)

        result["status"] = "OK"
        result["file"] = out_path
        result["detail"] += f"{len(output.splitlines())} lines saved"
        return result

    except NetmikoAuthenticationException:
        result["detail"] += "authentication failed"
    except NetmikoTimeoutException:
        result["detail"] += "connection timed out"
    except Exception as e:
        result["detail"] += f"error: {e}"
    return result


# --- Report -----------------------------------------------------------------
def write_report(results):
    report_dir = os.path.join(BASE_DIR, "report")
    os.makedirs(report_dir, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M")
    path = os.path.join(report_dir, f"Config_Backup_Report_{stamp}.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Hostname", "IP Address", "Device Type",
                         "Status", "Detail", "Saved File"])
        for r in results:
            writer.writerow([r["hostname"], r["ip"], r["device_type"],
                             r["status"], r["detail"], r["file"]])
    return path


# --- Credentials ------------------------------------------------------------
def get_credentials():
    """From environment if set, otherwise prompt. Never written to disk."""
    username = os.environ.get("NET_USER")
    password = os.environ.get("NET_PASS")
    secret = os.environ.get("NET_ENABLE", "")

    if not username:
        username = input("SSH username: ").strip()
    if not password:
        password = getpass.getpass("SSH password: ")
    if not secret:
        secret = getpass.getpass("Enable secret (blank if none): ")

    return {"username": username, "password": password, "secret": secret}


# --- Main -------------------------------------------------------------------
def main():
    print("=" * 60)
    print(" Config Backup Runner — Network Engineer Toolkit")
    print("=" * 60)

    devices = load_devices()
    if not devices:
        print("No devices found in devices.xlsx. Nothing to do.")
        return
    print(f"Loaded {len(devices)} device(s) from {os.path.basename(EXCEL_FILE)}.")

    creds = get_credentials()

    today = datetime.now().strftime("%Y-%m-%d")
    out_dir = os.path.join(BASE_DIR, "backups", today)
    os.makedirs(out_dir, exist_ok=True)
    print(f"Backups -> {out_dir}")
    print("-" * 60)

    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {
            pool.submit(backup_device, d, creds, out_dir): d for d in devices
        }
        for fut in as_completed(futures):
            r = fut.result()
            mark = "OK  " if r["status"] == "OK" else "FAIL"
            print(f"[{mark}] {r['hostname']:<20} {r['ip']:<16} {r['detail']}")
            results.append(r)

    ok = sum(1 for r in results if r["status"] == "OK")
    failed = len(results) - ok
    print("-" * 60)
    print(f"Done: {ok} OK, {failed} failed.")
    report_path = write_report(results)
    print(f"Report -> {report_path}")


if __name__ == "__main__":
    main()
