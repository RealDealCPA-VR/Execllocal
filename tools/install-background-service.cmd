@echo off
setlocal
REM ============================================================
REM  ExcelLocal background pane server (Windows)
REM  - starts hidden at every login (Startup folder, no admin needed)
REM  - serves the pane + LLM bridges on http://127.0.0.1:3000
REM  Undo: tools\uninstall-background-service.cmd
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
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":3000 .*LISTENING"') do taskkill /F /PID %%p >nul 2>nul

echo Writing hidden launcher to the Startup folder...
> "%REPO%\tools\serve-hidden.vbs" echo CreateObject^("WScript.Shell"^).Run "cmd /c cd /d ""%REPO%"" && node server\serve.js", 0, False
copy /y "%REPO%\tools\serve-hidden.vbs" "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ExcelLocalPaneServer.vbs" >nul

echo Starting it now...
wscript.exe "%REPO%\tools\serve-hidden.vbs"

echo Waiting for the pane server...
set /a tries=0
:waitloop
timeout /t 2 /nobreak >nul
set /a tries+=1
curl.exe -s -o NUL -w "%%{http_code}" http://127.0.0.1:3000/taskpane.html | find "200" >nul
if errorlevel 1 (
  if %tries% lss 10 goto waitloop
  echo Pane server did not come up on port 3000.
  pause & exit /b 1
)
echo Pane server is live on http://127.0.0.1:3000 ^(hidden - no window^).
echo Reload the ExcelLocal pane in Excel and you are set.
pause
