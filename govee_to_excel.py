#!/usr/bin/env python3
import os
import sys
import datetime as dt
from zoneinfo import ZoneInfo
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

API_KEY = os.getenv("GOVEE_API_KEY")
if not API_KEY:
    sys.exit("Set GOVEE_API_KEY in your environment.")

# Your two devices
DEVICES = [
    {"name": "Lab Fridge 1", "sku": "H5111", "id": "FE:32:D0:03:81:C6:31:81"},
    {"name": "Lab Fridge 2", "sku": "H5111", "id": "09:EC:D0:03:80:86:66:33"},
]

UNIT = os.getenv("GOVEE_TEMP_UNIT", "F").upper()  # "F" or "C"
XLSX_PATH = "govee_temps.xlsx"
SHEET_NAME = "readings"

def c_to_f(c): return (c * 9/5) + 32
def f_to_c(f): return (f - 32) * 5/9

def ensure_wb(path):
    try:
        wb = load_workbook(path)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.create_sheet(SHEET_NAME)
        if ws.max_row == 1 and ws.cell(1,1).value is None:
            ws.append(["timestamp","device_name","temp_f","temp_c"])
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(["timestamp","device_name","temp_f","temp_c"])
    return wb, ws

def autosize(ws):
    for col in range(1, ws.max_column + 1):
        width = max(len(str(c.value)) if c.value is not None else 0 for c in ws.iter_cols(min_col=col, max_col=col, values_only=False)[0])
        ws.column_dimensions[get_column_letter(col)].width = min(width + 2, 60)

def read_temp(session, sku, device_id):
    r = session.post(
        "https://openapi.api.govee.com/router/api/v1/device/state",
        headers={"Content-Type":"application/json","Govee-API-Key":API_KEY},
        json={"requestId":"uuid","payload":{"sku":sku,"device":device_id}},
        timeout=15
    )
    r.raise_for_status()
    caps = r.json().get("payload", {}).get("capabilities", [])
    val = next((c.get("state", {}).get("value") for c in caps if c.get("instance")=="sensorTemperature"), None)
    if val is None:
        raise RuntimeError("No temperature in response")
    v = float(val)
    if UNIT == "F":
        temp_f, temp_c = v, round(f_to_c(v), 2)
    else:
        temp_c, temp_f = v, round(c_to_f(v), 2)
    return round(temp_f, 2), round(temp_c, 2)

def main():
    ts = dt.datetime.now(ZoneInfo("America/Chicago")).isoformat(timespec="seconds")
    wb, ws = ensure_wb(XLSX_PATH)
    with requests.Session() as s:
        for d in DEVICES:
            tf, tc = read_temp(s, d["sku"], d["id"])
            ws.append([ts, d["name"], tf, tc])
    autosize(ws)
    wb.save(XLSX_PATH)
    print(f"Logged {len(DEVICES)} rows to {XLSX_PATH}")

if __name__ == "__main__":
    main()
