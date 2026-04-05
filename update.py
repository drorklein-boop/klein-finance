#!/usr/bin/env python3
import sys
from pathlib import Path
from openpyxl import load_workbook

EXCEL = Path(__file__).parent / "\u05de\u05d0\u05d6\u05df_\u05e7\u05dc\u05d9\u05d9\u05df.xlsm"

print("Testing Excel write...")
print("Excel path:", EXCEL)
print("Excel exists:", EXCEL.exists())
print("")

try:
    wb = load_workbook(EXCEL, keep_vba=True)
    print("Sheets found:", [s for s in wb.sheetnames if "RSU" in s or "ALIGN" in s])
    
    if "ALIGN RSU" in wb.sheetnames:
        ws = wb["ALIGN RSU"]
        print("H13 before:", ws.cell(row=13, column=8).value)
        print("H14 before:", ws.cell(row=14, column=8).value)
        ws.cell(row=13, column=8).value = 170600
        ws.cell(row=14, column=8).value = 187148
        wb.save(EXCEL)
        print("")
        print("SUCCESS! Written 170600 and 187148")
    else:
        print("ERROR: ALIGN RSU sheet not found!")
except Exception as e:
    print("ERROR:", e)

print("")
input("Press Enter to close...")
