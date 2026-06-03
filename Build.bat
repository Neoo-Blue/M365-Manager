@echo off
title M365Admin - Build to EXE
cd /d "%~dp0app"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0app\Build.ps1"
