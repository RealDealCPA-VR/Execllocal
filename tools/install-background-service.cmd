@echo off
setlocal
REM ============================================================
REM  ExcelLocal background pane server (Windows)
REM  - starts hidden at every login, no terminal window
REM  - serves the pane + LLM bridges on http://127.0.0.1:3000
REM  Run this once; afterwards Excel just works. Uninstall:
REM  tools\uninstall-background-service.cmd
REM ============================================================

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required. Install the LTS from https://nodejs.org
  pause & exit /b 1
)

for %%i in ("%~dp0..\") do set "REPO=%%~fi"
echo Repo: %REPO%

if not exist "%REPO%\dist\taskpane.html" (
  echo Building the pane ^(first time, a minute or two^)...
  pushd "%REPO%" && call npm install && call npm run build && popd
  if errorlevel 1 ( echo Build failed. & pause & exit /b 1 )
)

echo Stopping the dev server if it is running ^(frees port 3000^)...
pushd "%REPO%" && call npm run stop >nul 2>nul & popd

echo Writing hidden launcher...
> "%REPO%\tools\serve-hidden.vbs" echo CreateObject^("WScript.Shell"^).Run "cmd /c cd /d ""%REPO%"" && node server\serve.js", 0, False

echo Creating scheduled task ^(runs hidden at every login^)...
schtasks /create /f /tn "ExcelLocalPaneServer" /tr "wscript.exe \"%REPO%\tools\serve-hidden.vbs\"" /sc onlogon
if errorlevel 1 ( echo Could not create the scheduled task. & pause & exit /b 1 )

echo Starting it now...
schtasks /run /tn "ExcelLocalPaneServer"

echo Waiting for the pane server...
set /a tries=0
:waitloop
timeout /t 2 /nobreak >nul
set /a tries+=1
curl.exe -s -o NUL -w "%%{http_code}" http://127.0.0.1:3000/taskpane.html | find "200" >nul
if errorlevel 1 (
  if %tries% lss 10 goto waitloop
  echo Pane server did not come up on port 3000. Check Task Scheduler for "ExcelLocalPaneServer".
  pause & exit /b 1
)
echo Pane server is live on http://127.0.0.1:3000 ^(hidden - no window^).
echo Reload the ExcelLocal pane in Excel ^(or restart Excel^) and you are set.
pause
