@echo off
REM Run this ONCE on a Windows PC that has Python 3.10+ installed (python.org, tick "Add to PATH").
REM It creates dist\ToolToEmail.exe - a single file that runs on any Windows PC without installing anything.
py -m pip install --upgrade pip
py -m pip install pyinstaller openpyxl
py -m PyInstaller --onefile --windowed --name ToolToEmail app.py
echo.
echo Done. Copy dist\ToolToEmail.exe (and optionally config.ini) to any computer.
pause
