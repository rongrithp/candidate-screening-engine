#!/usr/bin/env python3
"""
PyInstaller Build Helper (`build_exe.py`)
Bundles `screener_core.py` into a single standalone executable (`screener_core.exe`).
"""

import sys
import os
import subprocess
import shutil
from pathlib import Path

# Force UTF-8 stdout encoding for Windows CLI compatibility
os.environ["PYTHONUNBUFFERED"] = "1"
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def build():
    workspace_root = Path(__file__).resolve().parent
    entry_script = workspace_root / "screener_core.py"
    
    if not entry_script.exists():
        print(f"Error: Entry script '{entry_script}' not found.")
        sys.exit(1)

    # Clean existing build artifacts
    for item in ["build", "dist", "screener_core.spec", "screener_core.exe"]:
        item_path = workspace_root / item
        if item_path.is_dir():
            shutil.rmtree(item_path, ignore_errors=True)
        elif item_path.is_file():
            try:
                item_path.unlink()
            except Exception:
                pass

    print("============================================================")
    print("Building Single Standalone Executable with PyInstaller")
    print("Target Script: screener_core.py")
    print("============================================================")
    
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name", "screener_core",
        str(entry_script)
    ]
    
    print(f"Command: {' '.join(cmd)}\n")
    res = subprocess.run(cmd, cwd=str(workspace_root))
    
    if res.returncode == 0:
        dist_exe = workspace_root / "dist" / "screener_core.exe"
        if not dist_exe.exists():
            dist_exe = workspace_root / "dist" / "screener_core"
            
        root_exe = workspace_root / "screener_core.exe"
        if dist_exe.exists() and dist_exe.is_file():
            shutil.copy2(dist_exe, root_exe)
            print("\n============================================================")
            print(f"[OK] Standalone Executable created successfully: {root_exe}")
            print("============================================================")
        else:
            print(f"\n[Warning] Built file not found at: {dist_exe}")
    else:
        print("\n[Error] PyInstaller build failed.")
        sys.exit(res.returncode)

if __name__ == "__main__":
    build()
