@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py docx_tables_to_excel_app.py
) else (
    python docx_tables_to_excel_app.py
)
pause
