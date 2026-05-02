@echo off
if not "%1"=="hidden" (
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\" hidden' -WindowStyle Hidden"
    exit /b
)

set DIR=%~dp0
cd /d "%DIR%"
"%DIR%python\pythonw.exe" "%DIR%mbot_manager.py"
