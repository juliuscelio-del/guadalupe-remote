@echo off
cd /d "%~dp0"
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name "Guadalupe NHS Remote" --icon icon.ico --hidden-import win32com.client --hidden-import pythoncom Remote.py
echo.
echo Done. Your app is in the "dist" folder: Guadalupe NHS Remote.exe
pause
