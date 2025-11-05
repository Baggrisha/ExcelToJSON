
# RU
## • [🇺🇸 EN](#EN)

# 📄 Excel → JSON Converter

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7+-blue?style=for-the-badge&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

**Профессиональный инструмент для поиска и извлечения данных из Excel файлов с возможностью сохранения результатов в JSON**

[🚀 Быстрый старт](#-быстрый-старт) • [📦 Скачать релиз](https://github.com/baggrisha/ExcelToJSON/releases/latest) • [📋 Возможности](#-возможности) • [💻 Установка](#-установка) • [🎯 Использование](#-использование) • [🇺🇸 EN](#EN)

</div>

---

## 🎯 Возможности

### ✨ Основные функции

* **🖥️ Графический интерфейс** – Удобный GUI на Tkinter  
* **📦 Пакетная обработка Excel** – Одновременная работа с несколькими файлами  
* **🔍 Поиск по содержимому** – По словам, столбцам и строкам  
* **📊 Сохранение результатов** – Экспорт в формат JSON  
* **🎨 Гибкие режимы поиска** – По названиям/индексам столбцов и строк, по двум колонкам  

---

## 🚀 Быстрый старт

> 💾 **Готовые сборки для macOS и Windows:**  
> [⬇️ Скачать последнюю версию →](https://github.com/baggrisha/ExcelToJSON/releases/latest)

---

### Установка из исходников

```bash
# 1. Клонируйте репозиторий
git clone https://github.com/baggrisha/ExcelToJSON.git
cd ExcelToJSON

# 2. Установите зависимости
pip install -r requirements.txt

# 3. Запустите программу
python XLStoJSON.py
````

---

## 📝 Поддерживаемые режимы

| Режим                                    | Описание                                                |
| ---------------------------------------- | ------------------------------------------------------- |
| **По слову**                             | Поиск слова или его вариаций в любом столбце            |
| **Достать весь текст из столбцов**       | Извлечение всех значений выбранного столбца по названию |
| **Достать весь текст из столбцов index** | По индексу столбца                                      |
| **Достать весь текст из строк**          | Поиск по строкам на содержание ключевого слова          |
| **Достать весь текст из строк index**    | По индексу строки                                       |
| **По двум столбцам**                     | Создание пары ключ-значение по названиям столбцов       |
| **По двум столбцам index**               | По индексам столбцов                                    |
| **Сохранить все данные**                 | Извлечение и сохранение всего содержимого Excel         |

---

## 💻 Установка

### Требования

* **Python 3.7+**
* **pandas**
* **openpyxl**
* **tkmacosx**
* **tkinter** (входит в стандартную поставку Python)

---

## 🎯 Использование

1. **Запустите программу**

   ```bash
   python XLStoJSON.py
   ```

2. **Выберите Excel файлы**

   * Нажмите **"Выбрать Excel"**
   * Укажите один или несколько `.xls/.xlsx` файлов

3. **Выберите режим**

   * Из выпадающего списка выберите нужный режим (по слову, по строкам и т.д.)

4. **Введите ключ или индекс**

   * Укажите слово или номер строки/столбца
   * Для режима "по двум столбцам" используйте оба поля

5. **Поиск и результаты**

   * Нажмите **"Поиск"**, программа покажет количество совпадений

6. **Сохранение**

   * Укажите папку для сохранения
   * Нажмите **"Сохранить в JSON"**

---

## 📋 Примеры использования

#### Поиск по слову

```
Слово: "Москва"
Результат: JSON с листами и значениями, где встречается "Москва"
```

#### Двумя столбцами

```
Столбец ключ: "ID"
Столбец значение: "Имя"
Результат: JSON с парами ключ-значение
```

#### Сохранение всех данных

```
Режим: "Сохранить все данные"
Результат: JSON со всеми столбцами и строками Excel
```

---

## 🏗️ Структура проекта

```
ExcelToJSON/
├── 📄 excel_to_json_gui.py      # Основной скрипт с GUI
├── 📋 requirements.txt          # Зависимости проекта
├── 📖 README.md                 # Документация (этот файл)
├── 📄 INSTRUCTION.txt           # Краткая инструкция
└── 🚀 run_converter.bat         # Быстрый запуск для Windows (локально)
```

---

## ⚙️ Технические детали

* **JSON**: UTF-8
* **Поддерживаемые форматы**: `.xls`, `.xlsx`
* **Выходной формат**: `.json`
* **Регистр**: не имеет значения (поиск нечувствителен к регистру)

---

## 🔧 Решение проблем

### ❓ Программа не запускается

Проверьте, что установлены зависимости:

```bash
python3 -m pip install pandas openpyxl pillow
```

### ❓ Не сохраняется JSON

Убедитесь, что выбрана папка для сохранения.

### ❓ Ошибка при открытии Excel

Проверьте, что файл не повреждён и имеет расширение `.xls` или `.xlsx`.

---

### ⚠️ Ошибка:

```
Traceback (most recent call last):
  File "XLStoJSON.py", line 5, in <module>
ModuleNotFoundError: No module named 'pandas'
```

#### **Причина:**

Не установлены библиотеки (`pandas`, `openpyxl`, `pillow`)
или PyInstaller не включил их в `.app` при сборке.

#### **Решение:**

##### 🧩 Если запускаете `.py`:

```bash
pip install pandas openpyxl pillow
```

##### 🍏 Если запускаете `.app` (скомпилированную версию):

Пересоберите приложение с зависимостями:

```bash
pyinstaller --onedir --windowed \
  --hidden-import pandas \
  --hidden-import tkmacosx \
  --hidden-import openpyxl \
  --hidden-import pillow \
  --icon=ico.png --clean XLStoJSON.py
```

После пересборки новая `.app` в папке `dist` будет работать на любом Mac.

---
# EN

## • [🇷🇺 RU](#RU)

# 📄 Excel → JSON Converter

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7+-blue?style=for-the-badge\&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

**Professional tool for searching and extracting data from Excel files with JSON export**

[🚀 Quick Start](#-quick-start) • [📦 Download Release](https://github.com/baggrisha/ExcelToJSON/releases/latest) • [📋 Features](#-features) • [💻 Installation](#-installation) • [🎯 Usage](#-usage) • [🇷🇺 RU](#RU)

</div>

---

## 🎯 Features

* **🖥️ Graphical Interface** – Tkinter-based GUI
* **📦 Batch Excel Processing** – Process multiple files simultaneously
* **🔍 Content Search** – By word, column, or row
* **📊 Save Results** – Export results to JSON
* **🎨 Flexible Modes** – Search by names/indexes or column pairs

---

## 🚀 Quick Start

> 💾 **Prebuilt apps for macOS and Windows:**
> [⬇️ Download latest release →](https://github.com/baggrisha/ExcelToJSON/releases/latest)

---

### Installation from source

```bash
# 1. Clone repository
git clone https://github.com/baggrisha/ExcelToJSON.git
cd ExcelToJSON

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run program
python XLStoJSON.py
```

---

## 📝 Supported Modes

| Mode                            | Description                                     |
| ------------------------------- | ----------------------------------------------- |
| **Search by word**              | Search for word or its variations in any column |
| **Extract entire column**       | Extract all values from a column by name        |
| **Extract column by index**     | Extract column by index                         |
| **Extract entire row**          | Search for a word in rows                       |
| **Extract row by index**        | By row index                                    |
| **Two-column mapping**          | Create key-value pairs by column names          |
| **Two-column mapping by index** | By column indexes                               |
| **Extract all data**            | Export entire Excel content                     |

---

## 💻 Installation

### Requirements

* **Python 3.7+**
* **pandas**
* **openpyxl**
* **tkmacosx**
* **tkinter** (included with Python)

---

## 🎯 Usage

1. **Run program**

   ```bash
   python XLStoJSON.py
   ```
2. **Select Excel files**

   * Click **"Select Excel"**
   * Choose one or more `.xls/.xlsx` files
3. **Select mode**

   * Choose search/extract mode from dropdown menu
4. **Enter key or index**

   * Input word or row/column index
   * For two-column modes, use both fields
5. **Search and view results**

   * Click **"Search"** to display match count
6. **Save to JSON**

   * Choose save folder → click **"Save to JSON"**

---

## 📋 Examples

#### Search by word

```
Word: "Moscow"
Result: JSON with sheets and values containing "Moscow"
```

#### Two-column mapping

```
Key column: "ID"
Value column: "Name"
Result: JSON with key-value pairs
```

#### Extract all data

```
Mode: "Extract all data"
Result: JSON with all Excel content
```

---

## 🏗️ Project Structure

```
ExcelToJSON/
├── 📄 excel_to_json_gui.py
├── 📋 requirements.txt
├── 📖 README.md
├── 📄 INSTRUCTION.txt
└── 🚀 run_converter.bat
```

---

## ⚙️ Technical Details

* **JSON**: UTF-8
* **Input formats**: `.xls`, `.xlsx`
* **Output format**: `.json`
* **Case-insensitive search**

---

## 🔧 Troubleshooting

### ❓ Program doesn’t start

Install dependencies:

```bash
pip install pandas openpyxl pillow
```

### ❓ JSON not saved

Check save folder is selected.

---

### ⚠️ Error:

```
Traceback (most recent call last):
  File "XLStoJSON.py", line 5, in <module>
ModuleNotFoundError: No module named 'pandas'
```

#### **Reason:**

Missing dependencies (`pandas`, `openpyxl`, `pillow`)
or PyInstaller didn’t include them in `.app`.

#### **Solution:**

##### 🧩 If running `.py`:

```bash
pip install pandas openpyxl pillow
```

##### 🍏 If running `.app` on macOS:

Rebuild with dependencies:

```bash
pyinstaller --onedir --windowed \
  --hidden-import pandas \
  --hidden-import tkmacosx \
  --hidden-import openpyxl \
  --hidden-import pillow \
  --icon=ico.png --clean XLStoJSON.py
```

