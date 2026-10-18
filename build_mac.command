#!/bin/bash
# Run this ON A MAC (double-click, or: bash build_mac.command). Needs Python 3.10+ from python.org.
cd "$(dirname "$0")"
python3 -m pip install --upgrade pip
python3 -m pip install pyinstaller openpyxl
python3 -m PyInstaller --windowed --noconfirm --name ToolToEmail app.py
codesign --force --deep --sign - dist/ToolToEmail.app
echo
echo "Xong! Ung dung nam o: $(pwd)/dist/ToolToEmail.app"
open dist
