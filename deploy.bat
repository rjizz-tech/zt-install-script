@echo off
REM This batch file will run the SD-WAN deployment script with the necessary permissions.

REM Prompt the user to ensure they run this as an administrator if they haven't already.
echo Make sure you have right-clicked this .bat file and selected "Run as administrator".
echo.
pause
echo.

echo Launching the Installation and Configuration Script...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0deploy-sd-wan.ps1"

echo.
echo The script has finished. You can review any messages above.
pause
@echo on
