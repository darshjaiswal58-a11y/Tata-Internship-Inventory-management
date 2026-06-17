@echo off
setlocal
cd /d "%~dp0"
if not exist logs mkdir logs
start "Tata Inventory Server" /min "%ComSpec%" /c ""%~dp0python\python.exe" "%~dp0app.py" 1>>"%~dp0logs\server.out.log" 2>>"%~dp0logs\server.err.log""
timeout /t 5 /nobreak >nul
start http://127.0.0.1:8000
