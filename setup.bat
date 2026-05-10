@echo off
setlocal EnableDelayedExpansion

:: ============================================================
::  mBot Manager — Portable Setup
::  Run once to download Python + install dependencies.
::  Internet connection required.
:: ============================================================

set PYTHON_VER=3.11.9
set PYTHON_ZIP=python-3.11.9-embed-amd64.zip
set PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VER%/%PYTHON_ZIP%
set GETPIP_URL=https://bootstrap.pypa.io/get-pip.py
set DIR=%~dp0
set PYDIR=%DIR%python

echo.
echo ============================================================
echo  mBot Manager ^| Portable Setup
echo ============================================================
echo.

:: ── Already set up? ─────────────────────────────────────────
if exist "%PYDIR%\python.exe" (
    if exist "%PYDIR%\Lib\site-packages\PyQt6" (
        echo [OK] Already set up. Run launch.bat to start.
        pause & exit /b 0
    )
)

:: ── Download Python embeddable ───────────────────────────────
if not exist "%PYDIR%\python.exe" (
    echo [1/4] Downloading Python %PYTHON_VER% embeddable...
    if not exist "%PYDIR%" mkdir "%PYDIR%"

    :: Try PowerShell download
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol='Tls12,Tls13'; (New-Object Net.WebClient).DownloadFile('%PYTHON_URL%', '%DIR%%PYTHON_ZIP%')}" 2>nul
    if not exist "%DIR%%PYTHON_ZIP%" (
        echo [ERROR] Download failed. Check internet connection.
        pause & exit /b 1
    )

    echo [1/4] Extracting Python...
    powershell -Command "Expand-Archive -Path '%DIR%%PYTHON_ZIP%' -DestinationPath '%PYDIR%' -Force"
    del /q "%DIR%%PYTHON_ZIP%"

    if not exist "%PYDIR%\python.exe" (
        echo [ERROR] Extraction failed.
        pause & exit /b 1
    )
)

:: ── Enable site-packages in embeddable Python ────────────────
echo [2/4] Configuring Python...

:: The ._pth file must include "Lib\site-packages" and uncomment "import site"
set PTH=%PYDIR%\python311._pth
if exist "%PTH%" (
    powershell -Command "(Get-Content '%PTH%') -replace '#import site','import site' | Set-Content '%PTH%'"
    :: Add site-packages line if missing
    powershell -Command "$c=Get-Content '%PTH%'; if ($c -notmatch 'Lib\\\\site-packages') { Add-Content '%PTH%' 'Lib\site-packages' }"
)

:: ── Download get-pip ──────────────────────────────────────────
if not exist "%PYDIR%\get-pip.py" (
    echo [2/4] Downloading pip...
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol='Tls12,Tls13'; (New-Object Net.WebClient).DownloadFile('%GETPIP_URL%', '%PYDIR%\get-pip.py')}" 2>nul
    if not exist "%PYDIR%\get-pip.py" (
        echo [ERROR] Could not download get-pip.py
        pause & exit /b 1
    )
)

:: ── Install pip ───────────────────────────────────────────────
if not exist "%PYDIR%\Scripts\pip.exe" (
    echo [2/4] Installing pip...
    "%PYDIR%\python.exe" "%PYDIR%\get-pip.py" --no-warn-script-location -q
    if errorlevel 1 (
        echo [ERROR] pip install failed.
        pause & exit /b 1
    )
)

:: ── Install dependencies ──────────────────────────────────────
echo [3/4] Installing dependencies (this may take a few minutes)...
"%PYDIR%\python.exe" -m pip install --no-warn-script-location -q ^
    PyQt6>=6.4.0 ^
    pywin32>=306 ^
    pywinauto>=0.6.8 ^
    uiautomation>=2.0.18 ^
    psutil>=7.2.2

if errorlevel 1 (
    echo [ERROR] Dependency installation failed.
    pause & exit /b 1
)

:: ── pywin32 post-install (registers COM DLLs) ─────────────────
echo [3/4] Running pywin32 post-install...
"%PYDIR%\python.exe" "%PYDIR%\Lib\site-packages\pywin32_system32" 2>nul
"%PYDIR%\python.exe" -c "import pywin32_bootstrap" 2>nul
:: Copy pywin32 system DLLs into python dir so they're found at runtime
for %%f in ("%PYDIR%\Lib\site-packages\pywin32_system32\*.dll") do (
    copy /y "%%f" "%PYDIR%\" >nul 2>&1
)

:: ── Done ─────────────────────────────────────────────────────
echo [4/4] Setup complete!
echo.
echo ============================================================
echo  Run launch.bat to start mBot Manager.
echo ============================================================
echo.
pause
