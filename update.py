#!/usr/bin/env python3
import sys, os, subprocess, time
from pathlib import Path
from openpyxl import load_workbook

EXCEL = Path(__file__).parent / "\u05de\u05d0\u05d6\u05df_\u05e7\u05dc\u05d9\u05d9\u05df.xlsm"

print("Closing Excel if open...")
os.system('taskkill /f /im excel.exe 2>nul')
time.sleep(2)

print("Opening Excel file...")
try:
    wb = load_workbook(EXCEL, keep_vba=True)
    print("Sheets with RSU:", [s for s in wb.sheetnames if "RSU" in s or "ALIGN" in s])
    
    if "ALIGN RSU" in wb.sheetnames:
        ws = wb["ALIGN RSU"]
        print("H13 before:", ws.cell(row=13, column=8).value)
        print("H14 before:", ws.cell(row=14, column=8).value)
        ws.cell(row=13, column=8).value = 170600
        ws.cell(row=14, column=8).value = 187148
        wb.save(EXCEL)
        print("SAVED! H13=170600, H14=187148")
    else:
        print("ERROR: ALIGN RSU sheet not found!")
except Exception as e:
    print("ERROR:", e)

print("")
print("Reopening Excel...")
os.startfile(str(EXCEL))
input("Press Enter to close...")
