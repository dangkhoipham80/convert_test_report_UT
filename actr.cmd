@echo off
REM Web UI: JS in web/ + python serve_ui.py. Or use "actr" in PowerShell after Register.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0TestReportConvert-Menu.ps1"
exit /b %errorlevel%
