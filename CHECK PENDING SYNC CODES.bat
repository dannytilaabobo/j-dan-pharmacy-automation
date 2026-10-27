@echo off
color 0B
cd /d "%~dp0"
set PY="C:\Users\DESKTOP-POS\AppData\Local\Python\bin\python.exe"
%PY% Check_Pending_Sync_Codes.py
pause
