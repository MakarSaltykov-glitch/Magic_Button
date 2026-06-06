@echo off
REM Magic Button - Quick Start Guide for Windows Users
REM This script shows setup instructions

chcp 65001 >nul
setlocal enabledelayedexpansion

echo.
echo ╔════════════════════════════════════════╗
echo ║   Magic Button - Быстрый Старт        ║
echo ║   Word ↔ Excel Конвертер Таблиц       ║
echo ╚════════════════════════════════════════╝
echo.

echo ШАГИ ДЛЯ УСТАНОВКИ И ЗАПУСКА:
echo.

echo 1️⃣  УСТАНОВКА PYTHON (если ещё не установлен)
echo    - Откройте https://www.python.org/downloads/
echo    - Скачайте Python 3.10 или выше
echo    - ⚠️  ПРИ УСТАНОВКЕ ОБЯЗАТЕЛЬНО:
echo       ✓ Поставьте галочку "Add Python to PATH"
echo       ✓ Поставьте галочку "Install tcl/tk"
echo    - Нажмите "Install Now"
echo.

echo 2️⃣  УСТАНОВКА ЗАВИСИМОСТЕЙ (первый раз)
echo    - Откройте папку с программой
echo    - Дважды нажмите: install_requirements.bat
echo    - Дождитесь "Successfully installed..."
echo.

echo 3️⃣  ЗАПУСК ПРОГРАММЫ
echo    - Дважды нажмите: run_docx_tables_app.bat
echo    - Откроется окно приложения Magic Button
echo    - Выберите нужный режим работы:
echo      📄 Word → Excel    (таблицы из Word в Excel)
echo      📊 Excel → Word    (листы из Excel в Word)
echo.

echo 4️⃣  ИСПОЛЬЗОВАНИЕ
echo    Word → Excel:
echo      1. Выберите папку с файлами .docx
echo      2. Выберите куда сохранить Excel файл
echo      3. Нажмите "Начать: Word → Excel"
echo.
echo    Excel → Word:
echo      1. Выберите файл .xlsx
echo      2. Выберите куда сохранить Word файл
echo      3. Нажмите "Начать: Excel → Word"
echo.

echo 5️⃣  СОЗДАНИЕ STANDALONE EXE (опционально)
echo    - Откройте Command Prompt в папке программы
echo    - Введите команду:
echo      build_exe.bat
echo    - После завершения откроется папка dist\
echo    - Там будет файл MagicButton.exe
echo    - Скопируйте его куда хотите - он будет работать!
echo.

echo ════════════════════════════════════════
echo  📖 Для подробной справки откройте:
echo     USAGE.md (в текстовом редакторе)
echo.
echo  ❓ Возникли проблемы?
echo     1. Убедитесь что Python установлен правильно
echo     2. Попробуйте запустить install_requirements.bat снова
echo     3. Проверьте что .docx и .xlsx файлы не открыты в других программах
echo ════════════════════════════════════════
echo.

pause
