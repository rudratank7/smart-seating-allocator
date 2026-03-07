@echo off
REM Run helper for Windows (cmd)
REM Usage: run_project.bat  (double-click or run from cmd)

SETLOCAL
SET "PROJECT_ROOT=%~dp0"
echo Project root: %PROJECT_ROOT%

REM Invoke the PowerShell helper which ensures Python 3.11 is installed and runs the project
powershell -ExecutionPolicy Bypass -NoProfile -File "%PROJECT_ROOT%run_project.ps1"

ENDLOCAL
