@echo off
echo Stopping and removing the ExcelLocal pane server task...
schtasks /end /tn "ExcelLocalPaneServer" >nul 2>nul
schtasks /delete /f /tn "ExcelLocalPaneServer"
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":3000 .*LISTENING"') do taskkill /F /PID %%p >nul 2>nul
echo Removed. To use the add-in again, run: npm start
pause
