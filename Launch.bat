@echo off
title M365 Administration Tool
cd /d "%~dp0"
set MSAL_BROKER_ENABLED=0
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0Main.ps1"
pause
