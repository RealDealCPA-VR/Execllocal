@echo off
REM ExcelLocal one-click launcher (Windows).
REM Starts the dev server with built-in LLM bridges and sideloads into Excel.

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required. Install it from https://nodejs.org and re-run.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Installing dependencies ^(first run only^)...
  call npm install
  if errorlevel 1 ( echo npm install failed. & pause & exit /b 1 )
)

REM Optional: point at a different LLM server, e.g.:
REM set VLLM_URL=http://localhost:8000

echo Starting ExcelLocal... Excel will open. Use Home ribbon - ExcelLocal.
call npm start
pause
