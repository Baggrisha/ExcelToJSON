
---

# RU
## • [🇺🇸 EN](#EN)

# 📄 Excel → JSON Converter

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7+-blue?style=for-the-badge\&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

**Профессиональный инструмент для поиска и извлечения данных из Excel файлов с возможностью сохранения результатов в JSON**

[🚀 Быстрый старт](#-быстрый-старт) • [📋 Возможности](#-возможности) • [💻 Установка](#-установка) • [🎯 Использование](#-использование) • [🇺🇸 EN](#EN)

</div>

---

## 🎯 Возможности

### ✨ Основные функции

* **🖥️ Графический интерфейс** – Удобный GUI на Tkinter
* **📦 Пакетная обработка Excel** – Одновременная работа с несколькими файлами
* **🔍 Поиск по содержимому** – По словам, столбцам и строкам
* **📊 Сохранение результатов** – Экспорт в формат JSON
* **🎨 Гибкие режимы поиска** – По названиям/индексам столбцов и строк, по двум колонкам

### 📝 Поддерживаемые режимы

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

## 🚀 Быстрый старт

### Установка

```bash
# 1. Клонируйте репозиторий
git clone https://github.com/baggrisha/ExcelToJSON.git
cd ExcelToJSON

# 2. Установите зависимости
pip install -r requirements.txt

# 3. Запустите программу
python excel_to_json_converter.py
```

### Windows (Быстрый запуск)

Для Windows пользователей доступен батник `run_converter.bat` (локально):

1. Дважды кликните по `run_converter.bat`
2. Программа автоматически установит зависимости и откроется

---

## 💻 Установка

### Требования

* **Python 3.7+**
* **pandas**
* **openpyxl**
* **tkinter** (входит в стандартную поставку Python)

### Установка зависимостей

```bash
# Автоматическая установка
pip install -r requirements.txt

# Или вручную
pip install pandas openpyxl
```

---

## 🎯 Использование

### Пошаговая инструкция

1. **Запустите программу**

   ```bash
   python excel_to_json_converter.py
   ```

2. **Выберите Excel файлы**

   * Нажмите **"Выбрать Excel"**
   * Выберите один или несколько `.xls/.xlsx` файлов

3. **Выберите режим поиска**

   * Выберите нужный режим из выпадающего меню (по слову, столбцу, строке и т.д.)

4. **Укажите ключ или индекс**

   * Введите слово или индекс строки/столбца в поле ввода
   * При необходимости используйте второе поле для режима "по двум столбцам"

5. **Поиск и просмотр результатов**

   * Нажмите **"Поиск"**, программа покажет количество найденных совпадений

6. **Сохранение в JSON**

   * Нажмите **"Выбрать место для сохранения"** и укажите директорию
   * Нажмите **"Сохранить в JSON"**, файлы будут созданы с соответствующими именами

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

### Форматирование по умолчанию

* **JSON**: UTF-8
* **Структура**: ключи — листы или значения столбцов, массивы — данные
* **Поиск**: поддерживаются вариации слова (нижний/верхний регистр, транслит)

### Поддерживаемые форматы

* **Входные**: `.xls`, `.xlsx`
* **Выходные**: `.json`

---

## 🔧 Решение проблем

### Частые вопросы

**Q: Программа не запускается**

```bash
python --version
pip install pandas openpyxl
```

**Q: Не сохраняется JSON**

* Проверьте, выбрана ли папка для сохранения

**Q: Ошибка при открытии Excel**

* Проверьте, что файл не поврежден и поддерживается форматом `.xls/.xlsx`

### Логи и отладка

* Программа выводит сообщения об ошибках через интерфейс Tkinter

---

# EN
## •[🇷🇺 RU](#RU)

# 📄 Excel → JSON Converter

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7+-blue?style=for-the-badge\&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange?style=for-the-badge)

**Professional tool for searching and extracting data from Excel files with JSON export**

[🚀 Quick Start](#-quick-start) • [📋 Features](#-features) • [💻 Installation](#-installation) • [🎯 Usage](#-usage) • [🇷🇺 RU](#RU)

</div>

---

## 🎯 Features

### ✨ Main Functions

* **🖥️ Graphical Interface** – Convenient GUI using Tkinter
* **📦 Batch Excel Processing** – Work with multiple files at once
* **🔍 Content Search** – By word, column, or row
* **📊 Save Results** – Export to JSON format
* **🎨 Flexible Modes** – By name/index of columns/rows, two-column pairs

### 📝 Supported Modes

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

## 🚀 Quick Start

### Installation

```bash
# 1. Clone repository
git clone https://github.com/baggrisha/ExcelToJSON.git
cd ExcelToJSON

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run program
python excel_to_json_converter.py
```

### Windows (Quick Launch)

1. Double-click `run_converter.bat`
2. Dependencies will be installed automatically, program will start

---

## 💻 Installation

### Requirements

* **Python 3.7+**
* **pandas**
* **openpyxl**
* **tkinter** (included with Python)

### Install Dependencies

```bash
# Automatic
pip install -r requirements.txt

# Or manually
pip install pandas openpyxl
```

---

## 🎯 Usage

### Step-by-step guide

1. **Run program**

   ```bash
   python excel_to_json_converter.py
   ```

2. **Select Excel files**

   * Click **"Select Excel"**
   * Choose one or more `.xls/.xlsx` files

3. **Select mode**

   * Choose search/extract mode from dropdown menu

4. **Enter key or index**

   * Input word or row/column index
   * For two-column modes, use second input

5. **Search and view results**

   * Click **"Search"** to see the number of matches

6. **Save to JSON**

   * Click **"Select save folder"**
   * Click **"Save to JSON"**, files will be created with corresponding names

---

## 📋 Usage Examples

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
├── 📄 excel_to_json_gui.py      # Main GUI script
├── 📋 requirements.txt          # Dependencies
├── 📖 README.md                 # Documentation (this file)
├── 📄 INSTRUCTION.txt           # Quick guide
└── 🚀 run_converter.bat         # Quick launcher for Windows (local)
```

---

## ⚙️ Technical Details

### Default Formatting

* **JSON**: UTF-8
* **Structure**: keys — sheets or column names, arrays — data
* **Search**: supports word variations (lowercase/uppercase/translit)

### Supported Formats

* **Input**: `.xls`, `.xlsx`
* **Output**: `.json`

---

## 🔧 Troubleshooting

### Common Issues

**Q: Program doesn’t start**

```bash
python --version
pip install pandas openpyxl
```

**Q: JSON not saved**

* Make sure a save folder is selected

**Q: Excel file cannot open**

* Verify file is `.xls/.xlsx` and not corrupted

### Logs and Debugging

* Program displays detailed error messages via Tkinter interface

---

