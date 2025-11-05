

# 🇷🇺 RU

## • [🇺🇸 EN](#-EN)

# 📦 EXCEL → JSON CONVERTER

**Быстрый старт и инструкция по запуску**

---

## 🚀 СКАЧАТЬ ГОТОВОЕ ПРИЛОЖЕНИЕ

* 🔗 **[Скачать последнюю версию (Releases)](https://github.com/baggrisha/ExcelToJSON/releases)**

  * macOS: `ExcelToJSON.app.zip` — распакуйте и дважды кликните
  * Windows: `ExcelToJSON.exe.zip` — распакуйте и запустите

> ⚠️ На macOS при первом запуске может появиться сообщение:
> *“Приложение загружено не из App Store”*
> → нажмите **ПКМ → Открыть → Открыть всё равно**

---

## 🧩 ЕСЛИ ПРЕДПОЧИТАЕТЕ ЗАПУСКАТЬ ЧЕРЕЗ PYTHON

### 📋 Установка вручную

1. Склонируйте репозиторий или скачайте `.zip`:

   ```bash
   git clone https://github.com/baggrisha/ExcelToJSON.git
   cd ExcelToJSON
   ```

2. Установите зависимости (один раз):

   ```bash
   pip install pandas openpyxl pillow
   ```

3. Запустите:

   ```bash
   python XLStoJSON.py
   ```

---

## 🎯 КАК ИСПОЛЬЗОВАТЬ

1. Нажмите **"Выбрать Excel"** и выберите один или несколько файлов
2. Нажмите **"Выбрать место для сохранения"** и укажите папку
3. Выберите режим поиска
4. При необходимости введите ключ, индекс или столбцы
5. Нажмите **"Поиск"**, чтобы проверить результаты
6. Нажмите **"Сохранить в JSON"**
7. ✅ Готово! Файлы JSON будут сохранены в выбранной папке

---

## ✨ ПОДДЕРЖИВАЕМЫЕ ФУНКЦИИ

✓ Поиск по слову с вариациями
✓ Поиск по столбцу или индексу
✓ Поиск по строкам или индексам
✓ По двум столбцам одновременно
✓ Извлечение всех данных
✓ Сохранение результата в JSON

---

## 🔧 ТРЕБОВАНИЯ

* Python 3.7 или выше
* pandas
* openpyxl
* pillow
* tkinter (входит в стандартную поставку Python)

---

## ❓ ЧАСТЫЕ ВОПРОСЫ

**В:** Можно ли обрабатывать несколько файлов сразу?
**О:** Да, просто выберите несколько файлов при выборе.

**В:** Какой формат сохраняется?
**О:** Все выбранные строки и столбцы сохраняются в JSON.

**В:** Можно ли искать по двум столбцам сразу?
**О:** Да, выберите режим “по двум столбцам”.

**В:** Можно ли удалить файлы из списка перед конвертацией?
**О:** Да, выберите файлы и нажмите **"Удалить выбранные"**.

---

## 📧 РЕШЕНИЕ ПРОБЛЕМ

**Если программа не запускается:**

1. Проверьте Python:

   ```bash
   python --version
   ```
2. Установите зависимости:

   ```bash
   pip install pandas openpyxl pillow
   ```
3. Проверьте, что файлы Excel доступны и не повреждены.

---

### 🧠 Ошибка:

```
Traceback (most recent call last):
  File "XLStoJSON.py", line 5, in <module>
ModuleNotFoundError: No module named 'pandas'
```

#### 💡 Решение:

🔹 **Windows:**

```bash
pip install pandas openpyxl pillow
```

🔹 **macOS:**

```bash
python3 -m pip install pandas openpyxl pillow
```

🔹 **Если ошибка в .app (Mac-приложении):**
пересоберите приложение с зависимостями:

```bash
pyinstaller --onedir --windowed \
  --hidden-import pandas \
  --hidden-import tkmacosx \
  --hidden-import openpyxl \
  --hidden-import pillow \
  --icon=ico.png --clean XLStoJSON.py
```

---

# 🇺🇸 EN

## • [🇷🇺 RU](#-RU)

# 📦 EXCEL → JSON CONVERTER

**Quick start & usage guide**

---

## 🚀 DOWNLOAD READY APP

* 🔗 **[Download latest version (Releases)](https://github.com/baggrisha/ExcelToJSON/releases)**

  * macOS: `ExcelToJSON.app.zip` — unzip & double-click
  * Windows: `ExcelToJSON.exe.zip` — unzip & run

> ⚠️ macOS users: if warning appears
> *“App is from an unidentified developer”*
> → right-click → Open → confirm

---

## 🧩 RUN VIA PYTHON

### 📋 Manual Installation

1. Clone or download the repository:

   ```bash
   git clone https://github.com/baggrisha/ExcelToJSON.git
   cd ExcelToJSON
   ```

2. Install dependencies:

   ```bash
   pip install pandas openpyxl pillow
   ```

3. Run:

   ```bash
   python XLStoJSON.py
   ```

---

## 🎯 HOW TO USE

1. Click **"Select Excel"** → choose one or more files
2. Click **"Select output folder"** → choose destination
3. Choose a search mode
4. Enter key, index, or column names if needed
5. Click **"Search"** to preview results
6. Click **"Save to JSON"**
7. ✅ Done! JSON files will appear in the selected folder

---

## ✨ SUPPORTED FEATURES

✓ Word search (case & variation insensitive)
✓ Column or row index search
✓ Two-column key-value mapping
✓ Extract all data
✓ Save results as JSON

---

## 🔧 REQUIREMENTS

* Python 3.7+
* pandas
* openpyxl
* pillow
* tkinter (included with Python)

---

## ❓ FAQ

**Q:** Can I process multiple files at once?
**A:** Yes, select several Excel files.

**Q:** What JSON format is used?
**A:** Standard UTF-8 JSON with lists and key-value mappings.

**Q:** Can I search by two columns?
**A:** Yes, select "two-column search" mode.

**Q:** Can I remove files before conversion?
**A:** Yes, select them and click **"Remove selected"**.

---

## 📧 TROUBLESHOOTING

**If the program doesn’t start:**

1. Check Python:

   ```bash
   python --version
   ```
2. Install dependencies:

   ```bash
   pip install pandas openpyxl pillow
   ```
3. Verify Excel files are valid.

---

### 🧠 Error:

```
Traceback (most recent call last):
  File "XLStoJSON.py", line 5, in <module>
ModuleNotFoundError: No module named 'pandas'
```

#### 💡 Solution:

```bash
pip install pandas openpyxl pillow
```


🔹 **If using compiled .app on macOS:**
rebuild with hidden imports:

```bash
pyinstaller --onedir --windowed \
  --hidden-import pandas \
  --hidden-import tkmacosx \
  --hidden-import openpyxl \
  --hidden-import pillow \
  --icon=ico.png --clean XLStoJSON.py
```

---

