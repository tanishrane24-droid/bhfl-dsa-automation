"""
Build executable using PyInstaller
Run: python build_exe.py
"""

import PyInstaller.__main__
import os
from pathlib import Path

# Create assets directory if it doesn't exist
assets_dir = Path('assets')
assets_dir.mkdir(exist_ok=True)

print("Building BHFL_DSA_AUTOMATION.exe...")
print("="*60)

PyInstaller.__main__.run([
    'main.py',
    '--name=BHFL_DSA_AUTOMATION',
    '--onefile',
    '--windowed',
    '--icon=assets/icon.ico',
    '--add-data=assets:assets',
    '--collect-all=pandas',
    '--collect-all=openpyxl',
    '--collect-all=win32com',
    '--distpath=./dist',
    '--buildpath=./build',
    '--specpath=./build',
    '--hidden-import=tkinter',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=win32com.client',
])

print("="*60)
print("✓ Executable created: dist/BHFL_DSA_AUTOMATION.exe")
print("="*60)
