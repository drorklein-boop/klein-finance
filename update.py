#!/usr/bin/env python3
import sys
from pathlib import Path
from openpyxl import load_workbook

EXCEL = Path(__file__).parent / "\u05de\u05d0\u05d6\u05df_\u05e7\u05dc\u05d9\u05d9\u05df.xlsm"

print("Opening:", EXCEL)
print("Exists:", EXCEL.exists())

wb = load_workbook(EXCEL, keep_vba=True)
print("Sheets:", wb.sheetnames)

if "ALIGN RSU" in wb.sheetnames:
    ws = wb["ALIGN RSU"]
    print("H13 before:", ws.cell(row=13, column=8).value)
    print("H14 before:", ws.cell(row=14, column=8).value)
    ws.cell(row=13, column=8).value = 170600
    ws.cell(row=14, column=8).value = 187148
    print("H13 after:", ws.cell(row=13, column=8).value)
    print("H14 after:", ws.cell(row=14, column=8).value)
    wb.save(EXCEL)
    print("SAVED!")
else:
    print("ALIGN RSU sheet NOT FOUND")

import os
os.startfile(str(EXCEL))
input("Press Enter to close...")
