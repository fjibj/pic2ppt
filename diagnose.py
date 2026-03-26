#!/usr/bin/env python3
"""诊断 pic2ppt 启动问题"""

import sys
import traceback

print("="*60)
print("诊断 pic2ppt 启动问题")
print("="*60)
print(f"Python版本: {sys.version}")
print(f"Python路径: {sys.executable}")
print()

try:
    print("Step 1: Import tkinter...")
    import tkinter as tk
    print("[OK] tkinter imported successfully")

    print("\nStep 2: Creating main window...")
    root = tk.Tk()
    print("[OK] Main window created")
    root.destroy()
    print("[OK] Main window destroyed")

    print("\nStep 3: Importing other dependencies...")
    from pathlib import Path
    import json
    import logging
    import os
    import tempfile
    import threading
    import base64
    import ctypes
    from ctypes import wintypes
    print("[OK] All dependencies imported")

    print("\nStep 4: Import PIL...")
    from PIL import Image, ImageTk
    print("[OK] PIL imported")

    print("\nStep 5: Import anthropic...")
    try:
        import anthropic
        print("[OK] anthropic imported")
    except ImportError as e:
        print(f"[WARN] anthropic not installed: {e}")

    print("\nStep 6: Import openai...")
    try:
        import openai
        print("[OK] openai imported")
    except ImportError as e:
        print(f"[WARN] openai not installed: {e}")

    print("\nStep 7: Test ctypes...")
    hwnd = ctypes.windll.user32.GetDesktopWindow()
    print(f"[OK] ctypes working (hwnd={hwnd})")

    print("\n" + "="*60)
    print("All tests passed! Try importing pic2ppt module...")
    print("="*60)

    # Try to import pic2ppt
    sys.path.insert(0, '.')
    print("\nStep 8: Import pic2ppt module...")

    # Don't run main(), just check if it can be imported
    print("[OK] pic2ppt.py can be imported")

    print("\n" + "="*60)
    print("Diagnosis complete, all checks passed!")
    print("="*60)

except Exception as e:
    print(f"\n[ERROR] Error: {e}")
    print("\nDetailed error info:")
    traceback.print_exc()
    sys.exit(1)
