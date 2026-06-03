@echo off
title M365 Administration Tool
set MSAL_BROKER_ENABLED=0
set M365ADMIN_ROOT=%~dp0app
cd /d "%~dp0app"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0app\Main.ps1"
pause
