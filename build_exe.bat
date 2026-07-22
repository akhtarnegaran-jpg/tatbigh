@echo off
chcp 65001 >nul
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --clean --onefile --windowed --name Tatbigh main.py
echo.
echo فایل اجرایی در پوشه dist ساخته شد.
pause
