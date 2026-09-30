@echo off
cd /d "%~dp0"
set "PAYROLL_LEGACY_DB=%~dp0payroll.db"
if exist "%~dp0dist\PayrollDesk\PayrollDesk.exe" (
    start "" "%~dp0dist\PayrollDesk\PayrollDesk.exe"
    exit /b
)
if exist "%~dp0.desktop-venv\Scripts\pythonw.exe" (
    start "" "%~dp0.desktop-venv\Scripts\pythonw.exe" "%~dp0desktop.py"
    exit /b
)
echo Please install PayrollDesk-Setup.exe or follow the developer setup in README.md.
pause
