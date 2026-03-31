@echo off
setlocal

cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin rights...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"

if %errorlevel% neq 0 (
    echo.
    echo Install failed.
    pause
    exit /b 1
)

echo.
pause
