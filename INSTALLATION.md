# 🚀 Установка Magic Button

## 👥 Для обычных пользователей (БЕЗ Python!)

### Windows
**Самый простой способ:**

1. ⬇️ [Скачайте MagicButton.exe](https://github.com/MakarSaltykov-glitch/Magic_Button/releases)
2. 📂 Поместите файл в любую папку (например, `C:\Program Files\`)
3. 🖱️ Дважды нажмите на `MagicButton.exe`
4. ✅ Готово!

**Больше ничего устанавливать не нужно!**

---

### macOS
**Самый простой способ:**

1. ⬇️ [Скачайте MagicButton.app.zip](https://github.com/MakarSaltykov-glitch/Magic_Button/releases)
2. 📂 Разархивируйте файл (двойной клик на ZIP)
3. 🚀 Перетащите `MagicButton.app` в папку `Applications`
4. 🖱️ Дважды нажмите на `MagicButton` в Applications
5. ✅ Готово!

**Больше ничего устанавливать не нужно!**

> ℹ️ Если система говорит "Can't open" - сделайте правый клик на app и выберите "Open"

---

## 👨‍💻 Для разработчиков (с Python)

### Требования
- Python 3.7+ установлен
- Git (опционально)

### Windows

```bash
# Скачайте или клонируйте репозиторий
git clone https://github.com/MakarSaltykov-glitch/Magic_Button.git
cd Magic_Button

# Установите зависимости (первый раз)
pip install -r requirements.txt

# Запустите программу
python docx_tables_to_excel_app.py
```

Или просто дважды нажмите `run_docx_tables_app.bat`

### macOS / Linux

```bash
# Скачайте или клонируйте репозиторий
git clone https://github.com/MakarSaltykov-glitch/Magic_Button.git
cd Magic_Button

# Установите зависимости (первый раз)
pip3 install -r requirements.txt

# Запустите программу
python3 docx_tables_to_excel_app.py
```

---

## 🔨 Для сборки собственных приложений

### Создать Windows EXE

```bash
pip install pyinstaller
python build_windows.py
```

Готовый файл: `dist/MagicButton.exe`

### Создать macOS APP

```bash
pip install pyinstaller
bash build_macos_installer.sh
```

Готовый файл: `dist/MagicButton.app`

### Создать macOS DMG (для распространения)

```bash
bash create_dmg.sh
```

Готовый файл: `MagicButton-installer.dmg`

---

## ❓ Часто задаваемые вопросы

**Q: Нужно ли устанавливать Python для использования приложения?**  
A: Нет! Скачайте готовое `MagicButton.exe` или `MagicButton.app` - они работают без Python.

**Q: Где найти готовые приложения?**  
A: Откройте [GitHub Releases](https://github.com/MakarSaltykov-glitch/Magic_Button/releases) и скачайте нужное.

**Q: Может ли я поделиться приложением?**  
A: Да! Скачайте готовый EXE или APP и делитесь с кем угодно. Никаких ограничений (MIT лицензия).

**Q: Мой антивирус говорит "опасный файл"**  
A: Это нормально для самосборных приложений. Вы можете проверить исходный код - он полностью открыт.

**Q: Как обновиться на новую версию?**  
A: Просто скачайте новый EXE/APP и замените старый.

---

## 🆘 Решение проблем

### Windows

**Проблема: "Windows защитил ваш компьютер"**
- Нажмите "Дополнительные сведения"
- Нажмите "Запустить всё равно"

**Проблема: Программа не запускается**
- Убедитесь что это правильный файл (MagicButton.exe)
- Скачайте снова с GitHub
- Проверьте место расположения файла (не должно быть спецсимволов в пути)

### macOS

**Проблема: "Can't open MagicButton.app"**
- Откройте Finder
- Найдите MagicButton.app
- Сделайте правый клик
- Выберите "Open"
- Нажмите "Open" в диалоге

**Проблема: "Permission denied" ошибка**
```bash
chmod +x /Applications/MagicButton.app/Contents/MacOS/MagicButton
```

---

## 📞 Помощь

- 📖 [Полное руководство](USAGE.md)
- 🐛 [Сообщить об ошибке](https://github.com/MakarSaltykov-glitch/Magic_Button/issues)
- 💬 [Задать вопрос](https://github.com/MakarSaltykov-glitch/Magic_Button/discussions)

---

**Готово! Наслаждайтесь Magic Button! 🎉**
