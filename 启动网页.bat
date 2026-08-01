@echo off
title Face Recognition Server - Close this window to STOP
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set INSIGHTFACE_HOME=%CD%\runtime\insightface_home

echo ==========================================
echo   Face Recognition Attendance - Web UI
echo   Browser will open automatically...
echo   CLOSE THIS WINDOW to stop the server.
echo ==========================================
start "" /min powershell -NoProfile -WindowStyle Hidden -Command "$c=New-Object Net.Sockets.TcpClient; $n=0; while($true){ try{$c.Connect('127.0.0.1',5000); break}catch{}; Start-Sleep -m 400; $n++; if($n -gt 75){break} }; Start-Process 'http://localhost:5000'"
runtime\python\python.exe web_show\web_app.py
echo.
echo Server stopped. You can close this window now.
pause
