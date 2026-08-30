@echo off
cd /d "%~dp0"
set PY="C:\Users\DESKTOP-POS\AppData\Local\Python\bin\python.exe"
start "Timekeeping App" cmd /k %PY% app.py
timeout /t 3 /nobreak >nul
start chrome "http://localhost:5000/"
ngrok http --domain=chalice-hut-clip.ngrok-free.dev 5000
