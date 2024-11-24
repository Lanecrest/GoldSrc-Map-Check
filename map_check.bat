@echo off
:Start
cls
echo Running PowerShell script...
powershell -ExecutionPolicy Bypass -File "map_check.ps1"

echo.
set /p userInput="Do you want to run the script again? (Y/N): "
if /I "%userInput%"=="Y" goto Start
echo Exiting...