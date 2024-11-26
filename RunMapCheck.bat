@echo off
:Start
cls
echo GoldSrc Map Check v1.3 (c) 2024 Lanecrest Tech
echo.
echo Running PowerShell script...
powershell -ExecutionPolicy Bypass -File "scripts/GenerateExcelFile.ps1"

echo.
set /p userInput="Do you want to run the script again? (Y/N): "
if /I "%userInput%"=="Y" goto Start
echo Exiting...