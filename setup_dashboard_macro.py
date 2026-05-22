# Klein Finance - One-time VBA setup
# Adds RunDashboard macro to the open workbook
import sys
from pathlib import Path

BASE = Path(__file__).parent

VBA_CODE = """
Sub RunDashboard()
    Dim wsh As Object
    If Dir("C:\\KleinFinance\\launcher_dashboard.py") = "" Then
        MsgBox "launcher_dashboard.py not found in C:\\KleinFinance\\", vbCritical
        Exit Sub
    End If
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /k python C:\\KleinFinance\\launcher_dashboard.py", 1, False
    Set wsh = Nothing
End Sub
"""

def main():
    print("\n  Klein Finance - Dashboard Macro Setup")
    print("  ======================================")

    try:
        import xlwings as xw
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "xlwings", "--quiet"], check=True)
        import xlwings as xw

    app = xw.apps.active
    if not app:
        print("  ERROR: Excel is not open.")
        input("\n  Press Enter to close..."); return

    wb = next((b for b in app.books if b.name.lower().endswith(".xlsm")), None)
    if not wb:
        print("  ERROR: No .xlsm workbook open.")
        input("\n  Press Enter to close..."); return

    print(f"  Workbook: {wb.name}")

    try:
        vbc = wb.api.VBProject.VBComponents
    except Exception as e:
        print(f"\n  ERROR: Cannot access VBA project.")
        print(f"  You need to enable \'Trust access to VBA project object model\':")
        print(f"  Excel → File → Options → Trust Center → Trust Center Settings")
        print(f"  → Macro Settings → check \'Trust access to VBA project object model\'")
        print(f"  Then run this script again.")
        input("\n  Press Enter to close..."); return

    # Check if module already exists
    existing = None
    for i in range(1, vbc.Count + 1):
        comp = vbc.Item(i)
        if comp.Name == "DashboardLauncher":
            existing = comp
            break

    if existing:
        existing.CodeModule.DeleteLines(1, existing.CodeModule.CountOfLines)
        existing.CodeModule.AddFromString(VBA_CODE)
        print("  RunDashboard macro updated.")
    else:
        new_mod = vbc.Add(1)  # 1 = standard module
        new_mod.Name = "DashboardLauncher"
        new_mod.CodeModule.AddFromString(VBA_CODE)
        print("  RunDashboard macro added successfully.")

    wb.save()
    print("  Workbook saved.")
    print("\n  Done. You can now add a button and assign it to RunDashboard.")
    input("\n  Press Enter to close...")

if __name__ == "__main__":
    main()
