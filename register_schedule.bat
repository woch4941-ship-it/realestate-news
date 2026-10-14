@echo off
rem Register a Windows scheduled task that collects news every day at 07:30.
rem Usage: register_schedule.bat [HH:MM]   (default 07:30)
set RUNTIME=%1
if "%RUNTIME%"=="" set RUNTIME=07:30
schtasks /Create /F /SC DAILY /ST %RUNTIME% /TN "RealEstateNewsDaily" /TR "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%~dp0collect.ps1\""
echo.
echo Registered: every day at %RUNTIME%  (remove with: schtasks /Delete /TN RealEstateNewsDaily /F)
pause
