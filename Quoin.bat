@echo off
setlocal
cd /d "%~dp0"
title Quoin Core

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed or is not available on PATH.
  echo Install Node.js LTS, then double-click this file again.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Installing Quoin dependencies. This only needs to happen the first time.
  call npm install
  if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 1
  )
)

echo Starting Quoin Core...
start "" "http://localhost:3000"
call npm run dev

pause
