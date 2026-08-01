@echo off
title Stop Face Recognition Server
echo ==========================================
echo   Stop Face Recognition Server (port 5000)
echo ==========================================
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
  echo Killing process PID %%a ...
  taskkill /F /PID %%a >nul 2>&1
)
echo.
echo If you saw "Killing process" above, the server has been stopped.
echo Otherwise, no server was running on port 5000.
echo.
pause
