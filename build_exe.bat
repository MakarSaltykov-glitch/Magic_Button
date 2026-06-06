@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -m pip install -r requirements.txt pyinstaller
    py -m PyInstaller --noconfirm --onefile --windowed --name MagicButton docx_tables_to_excel_app.py
) else (
    python -m pip install -r requirements.txt pyinstaller
    python -m PyInstaller --noconfirm --onefile --windowed --name MagicButton docx_tables_to_excel_app.py
)
echo.
echo Готово. Если сборка прошла успешно, файл находится здесь:
echo dist\MagicButton.exe
pause
