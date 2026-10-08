@echo off
rem Stops the HFG add-in and removes it from Excel on this computer.
setlocal
cd /d "%~dp0"
call npm run stop
pause
