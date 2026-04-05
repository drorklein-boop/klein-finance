#!/usr/bin/env python3
"""
Klein Finance Launcher
======================
This file NEVER changes.
It downloads the latest update.py from GitHub and runs it.
The VBA button calls this file.
"""

import urllib.request, sys, subprocess, time
from pathlib import Path

BASE       = Path(__file__).parent
UPDATE_URL = "https://raw.githubusercontent.com/drorklein-boop/klein-finance/main/update.py"
LOCAL      = BASE / "update.py"

def main():
    print("\n  Klein Finance Launcher")
    print("  Checking for latest version...\n")
    try:
        # Force no-cache by adding timestamp to URL
        url = UPDATE_URL + "?t=" + str(int(time.time()))
        req = urllib.request.Request(url, headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
        with urllib.request.urlopen(req, timeout=15) as response:
            content = response.read()
        LOCAL.write_bytes(content)
        print("  \u2713 Latest version downloaded\n")
    except Exception as e:
        if LOCAL.exists():
            print(f"  \u26a0 Could not update ({e})")
            print("  Running existing version...\n")
        else:
            print(f"  \u2717 No internet and no local script.")
            input("\n  Press Enter to close...")
            sys.exit(1)

    subprocess.run([sys.executable, str(LOCAL)], cwd=str(BASE))

if __name__ == "__main__":
    main()
