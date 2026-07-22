@echo off
chcp 65001 >nul
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
echo.
echo نصب وابستگی‌ها تمام شد.
pause
