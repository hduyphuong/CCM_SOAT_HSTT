@echo off
chcp 65001 >nul
title Engine soat HSTT - 127.0.0.1:8765
cd /d "%~dp0"
echo Dang chay engine soat HSTT tai http://127.0.0.1:8765  (dong cua so nay de tat)
python engine\app.py 8765
pause
