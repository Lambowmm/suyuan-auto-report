@echo off
cd /d D:\
cd D:\报告生成工具\JYJ
call venv\Scripts\activate.bat
set WEASYPRINT_DLL_DIRECTORIES=C:\msys64\mingw64\bin
python generate_reports.py
pause
