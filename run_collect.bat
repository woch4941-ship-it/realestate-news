@echo off
rem Collect today's real-estate news, then open the viewer.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0collect.ps1"
start "" "%~dp0viewer.html"
