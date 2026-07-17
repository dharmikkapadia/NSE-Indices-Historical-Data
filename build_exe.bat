@echo off
REM ============================================================
REM   NSE Nifty Indices Downloader  -  build_exe.bat
REM   Builds a standalone Windows .exe using PyInstaller
REM
REM   Usage:  Double-click this file, or run from cmd:
REM             build_exe.bat
REM
REM   Prerequisite (one-time):  Python 3.9 or newer installed
REM   from https://www.python.org/downloads/  with the
REM   "Add Python to PATH" option ticked.
REM
REM   After build, the .exe lives at:  dist\NSEDataDownloader.exe
REM   That single .exe runs on any Windows PC -- no Python needed.
REM ============================================================

setlocal enabledelayedexpansion
echo.
echo === NSE Nifty Indices Downloader  -  Build Script ===
echo.

REM --- 1. Check Python is on PATH ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not on PATH.
    echo.
    echo   Install Python 3.9+ from https://www.python.org/downloads/
    echo   IMPORTANT: tick "Add python.exe to PATH" during install.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version') do set PYVER=%%v
echo Found: !PYVER!
echo.

REM --- 2. Install build dependencies ---
echo Installing build dependencies (pyinstaller, requests, customtkinter, darkdetect, openpyxl)...
python -m pip install --upgrade pip >nul
python -m pip install --upgrade pyinstaller requests certifi customtkinter darkdetect openpyxl
if errorlevel 1 (
    echo.
    echo [ERROR] Could not install dependencies.
    pause
    exit /b 1
)
echo.

REM --- 3. Clean previous builds ---
echo Cleaning previous builds...
if exist build           rmdir /s /q build
if exist dist            rmdir /s /q dist
if exist NSEDataDownloader.spec del /q NSEDataDownloader.spec
echo.

REM --- 4. Build the .exe ---
echo Building NSEDataDownloader.exe ...
echo (this can take 1-3 minutes the first time)
echo.

python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "NSEDataDownloader" ^
    --collect-all certifi ^
    --collect-all customtkinter ^
    --collect-all darkdetect ^
    --collect-all openpyxl ^
    nse_downloader.py

if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller build failed. See messages above.
    pause
    exit /b 1
)

REM --- 5. Optional cleanup of build intermediates ---
if exist build rmdir /s /q build
if exist NSEDataDownloader.spec del /q NSEDataDownloader.spec

echo.
echo ============================================================
echo   BUILD SUCCESS
echo.
echo   Your standalone executable is here:
echo     %CD%\dist\NSEDataDownloader.exe
echo.
echo   Copy that single file to any Windows PC and double-click
echo   to run -- no Python install required on that machine.
echo ============================================================
echo.
pause
endlocal
