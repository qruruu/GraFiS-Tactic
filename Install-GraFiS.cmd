@echo off
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0tools\Open-GraFiS.ps1" -Install
if errorlevel 1 pause
