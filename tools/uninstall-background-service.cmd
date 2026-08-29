@echo off
echo Removing the hidden ExcelLocal pane server...
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ExcelLocalPaneServer.vbs" >nul 2>nul
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":3000 .*LISTENING"') do taskkill /F /PID %%p >nul 2>nul
echo Removed. To use the add-in again, run: npm start
pause
