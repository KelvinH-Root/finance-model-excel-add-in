@echo off
rem Installs (first time only) and starts the HFG add-in in Excel. Double-click it, or run it from a terminal.
setlocal
cd /d "%~dp0"
title HFG add-in

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed. Install the LTS version from https://nodejs.org, then run this again.
  goto :fail
)
for /f "tokens=1 delims=v." %%v in ('node -v') do set NODEMAJOR=%%v
if %NODEMAJOR% LSS 20 (
  echo The add-in needs Node.js 20 or later. This computer has:
  node -v
  echo Install the LTS version from https://nodejs.org, then run this again.
  goto :fail
)

if not exist node_modules (
  echo Installing the add-in's tools. This happens once and takes a minute or two.
  call npm install
  if errorlevel 1 goto :fail
)

if not exist "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" (
  echo Installing a certificate so Excel trusts the add-in on this computer. This happens once.
  echo Windows asks whether to install it: choose Yes.
  call npm run certs
  if errorlevel 1 goto :fail
)

echo Starting the add-in. Excel opens with the HFG Model tab.
call npm start
if errorlevel 1 goto :fail
echo.
echo The add-in is running. In Excel, open the HFG Model tab and choose New model.
echo When you have finished, run stop-addin.cmd.
pause
exit /b 0

:fail
echo.
echo Something went wrong. Copy the messages above into the chat with Claude.
pause
exit /b 1
