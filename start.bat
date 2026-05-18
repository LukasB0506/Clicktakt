@echo off
cd /d "%~dp0"
pip install openpyxl PyQt6 >nul 2>&1
echo Starte Programm...
python takteinteilung_qt.py
if %errorlevel% neq 0 (
    echo FEHLER! Bitte Screenshot machen.
    pause
)