#!/usr/bin/env python3
"""Klein Finance Launcher - never changes"""
import urllib.request, sys, subprocess, time
from pathlib import Path

BASE       = Path(__file__).parent
UPDATE_URL = "https://raw.githubusercontent.com/drorklein-boop/klein-finance/main/main.py"
LOCAL      = BASE / "update.py"

def main():
    print("\n  Klein Finance Launcher")
    print("  Downloading latest version...\n")
    try:
        if LOCAL.exists():
            LOCAL.unlink()
        url = UPDATE_URL + "?t=" + str(int(time.time()))
        req = urllib.request.Request(url, headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
        with urllib.request.urlopen(req, timeout=15) as r:
            LOCAL.write_bytes(r.read())
        print("  \u2713 Latest version downloaded\n")
    except Exception as e:
        print(f"  \u26a0 Could not download: {e}")
        if not LOCAL.exists():
            input("\n  Press Enter to close...")
            sys.exit(1)
        print("  Running existing version...\n")
    subprocess.run([sys.executable, str(LOCAL)], cwd=str(BASE))

if __name__ == "__main__":
    main()
