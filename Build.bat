@echo off
title M365Admin - Build to EXE
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0Build.ps1"
