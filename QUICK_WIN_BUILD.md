# 🪟 Windows Build - Быстро для вашего заказчика

## Вариант 1: Я дам готовый EXE (ПРОЩЕ ВСЕГО)

Просто используй бинарный файл из GitHub Actions или я пришлю тебе готовый.

## Вариант 2: Собрать самому на Windows (если нужно)

```bash
# На Windows машине:
pip install -r requirements.txt
python build_windows.py
```

Результат: `dist\MagicButton.exe` (~90-100 MB)

## Вариант 3: Использовать py2exe (для Windows)

```bash
# На Windows:
pip install py2exe
python setup.py py2exe
```

---

**Что нужно заказчику Windows:**

1. Скачать MagicButton.exe
2. Двойной клик
3. Использовать приложение
4. Готово! ✅
