@echo off
title M365 Administration Tool
set MSAL_BROKER_ENABLED=0
set M365ADMIN_ROOT=%~dp0
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0Main.ps1"
pause
