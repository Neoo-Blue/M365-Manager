@echo off
title M365 Administration Tool
set MSAL_BROKER_ENABLED=0
set M365ADMIN_ROOT=%~dp0app
cd /d "%~dp0app"

REM Prefer PowerShell 7 (pwsh) when it is installed.
REM
REM Windows PowerShell 5.1 has no assembly isolation, so Exchange Online,
REM SharePoint Online and the Microsoft Graph SDK all fight over one copy
REM of shared dependencies like System.Text.Json. The usual casualty is
REM "The type initializer for 'Azure.Identity.AuthenticationRecord' threw
REM an exception" on the first Graph connect. PowerShell 7 loads the Graph
REM SDK in its own AssemblyLoadContext, which removes the fight entirely.
REM
REM 5.1 remains fully supported and is the automatic fallback.
REM Set M365ADMIN_FORCE_PS5=1 to pin this machine to 5.1.
REM
REM One thing to know when switching: PowerShell 7 does not read
REM Documents\WindowsPowerShell\Modules, so modules installed for
REM CurrentUser under 5.1 are invisible to it and the tool will install
REM its own copies on first run. AllUsers installs (what option 98 does)
REM are shared by both.
set "PSEXE="
if "%M365ADMIN_FORCE_PS5%"=="1" goto usePS5
for %%P in (pwsh.exe) do set "PSEXE=%%~$PATH:P"
if defined PSEXE goto run

:usePS5
set "PSEXE=powershell.exe"

:run
"%PSEXE%" -ExecutionPolicy Bypass -NoProfile -File "%~dp0app\Main.ps1"
pause
