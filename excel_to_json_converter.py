import json
import os
import sys
import tkinter as tk
import pandas as pd
from PIL import Image, ImageTk
from tkinter import filedialog, Listbox, END, messagebox

if sys.platform == "darwin":  # 'darwin' — это Mac
    from tkmacosx import Button
else:
    from tkinter import Button


# Класс для поиска данных в Excel файлах
class ExcelSearcher:
    def __init__(self, files):
        # Список Excel файлов
        self.files = files
        self.dfs = []  # Список DataFrame для каждого листа
        self.dfs_name = []  # Список названий листов
        self.load_files()  # Загружаем файлы при инициализации

    def load_files(self):
        """Загрузка всех выбранных Excel файлов и листов"""
        self.dfs.clear()
        self.dfs_name.clear()
        for file_name in self.files:
            try:
                excel_file = pd.ExcelFile(file_name)
                for sheet_name in excel_file.sheet_names:
                    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
                    df_sheet.columns = df_sheet.columns.astype(str).str.strip()  # Очистка названий колонок
                    self.dfs.append(df_sheet)
                    self.dfs_name.append(sheet_name)
            except Exception as e:
                print(f"Ошибка загрузки {file_name}: {e}")

    def generate_variations(self, word):
        """Генерация вариантов слова (регистры, английская раскладка)"""
        variants = set()
        variants.add(word.lower())
        variants.add(word.upper())
        variants.add(word.capitalize())
        # Преобразование русских букв в английские по клавиатуре
        eng_map = str.maketrans("фисвуапршолдьтщзйкыегмцчня", "abcdefghijklmnopqrstuvwxyz")
        variants.add(word.lower().translate(eng_map))
        return list(variants)

    def search_word(self, word):
        """Поиск по слову во всех листах"""
        variants = self.generate_variations(word)
        results = {}

        for sheet_name, df in zip(self.dfs_name, self.dfs):
            found_values = []
            for col in df.columns:
                for val in df[col].astype(str):
                    if any(v in val for v in variants):
                        found_values.append(val)
            if found_values:
                results[sheet_name] = found_values

        total_found = sum(len(v) for v in results.values())
        return results, total_found

    def search_column(self, column_name):
        """Поиск по названию столбца"""
        result = {}
        column_name_lower = column_name.lower()
        for df in self.dfs:
            matching_cols = [col for col in df.columns if column_name_lower in col.lower()]
            for col in matching_cols:
                for val in df[col].dropna():
                    result.setdefault(col, []).append(val)

        total_found = sum(len(v) for v in result.values())
        return result, total_found

    def search_column_by_index(self, column_index):
        """Поиск по индексу столбца"""
        result = {}
        for df in self.dfs:
            try:
                col = df.iloc[:, int(column_index)]
                col_name = df.columns[int(column_index)]
                for val in col.dropna():
                    result.setdefault(col_name, []).append(val)
            except (ValueError, IndexError):
                continue

        total_found = sum(len(v) for v in result.values())
        return result, total_found

    def search_rows(self, word):
        """Поиск по строкам с ключевым словом"""
        variants = self.generate_variations(word)
        result = {str(word): []}
        for df in self.dfs:
            for idx, row in df.iterrows():
                row_str = '; '.join(row.astype(str).tolist())
                if any(v in row_str for v in variants):
                    result[str(word)].append(row_str)

        total_found = sum(len(v) for v in result.values())
        return result, total_found

    def search_rows_by_index(self, row_index):
        """Поиск строки по индексу"""
        result = {str(row_index): []}
        for df in self.dfs:
            try:
                if str(row_index) in df.index.astype(str):
                    idx_match = df.index[df.index.astype(str) == str(row_index)][0]
                    row = df.loc[idx_match]
                    row_str = '; '.join(row.astype(str).tolist())
                    result[str(row_index)].append(row_str)
            except Exception:
                continue

        total_found = sum(len(v) for v in result.values())
        return result, total_found

    def search_two_columns(self, key_col, value_col):
        """Поиск по двум столбцам: ключ-значение"""
        key_col_lower = key_col.lower()
        value_col_lower = value_col.lower()
        result = {}
        for df in self.dfs:
            matching_keys = [col for col in df.columns if key_col_lower in col.lower()]
            matching_values = [col for col in df.columns if value_col_lower in col.lower()]
            for k_col in matching_keys:
                for v_col in matching_values:
                    for k, v in zip(df[k_col].astype(str), df[v_col].astype(str)):
                        k_val = k if k and k != "nan" else "NaN"
                        v_val = v if v and v != "nan" else "NaN"
                        result.setdefault(k_val, []).append(v_val)

        total_found = sum(len(v) for v in result.values())
        return result, total_found

    def search_two_columns_by_index(self, key_col_index, value_col_index):
        """Поиск ключ-значение по индексам столбцов"""
        result = {}
        for df in self.dfs:
            try:
                k_col = df.iloc[:, int(key_col_index)]
                v_col = df.iloc[:, int(value_col_index)]
                for k, v in zip(k_col.astype(str), v_col.astype(str)):
                    k_val = k if k and k != "nan" else "NaN"
                    v_val = v if v and v != "nan" else "NaN"
                    result.setdefault(k_val, []).append(v_val)
            except (ValueError, IndexError):
                continue

        total_found = sum(len(v) for v in result.values())
        return result, total_found

    def get_all_data(self):
        """Извлечение всех данных из всех файлов"""
        result = {}
        for df in self.dfs:
            for col in df.columns:
                result.setdefault(col, []).extend(df[col].dropna().tolist())

        total_found = sum(len(v) for v in result.values())
        return result, total_found


# Tkinter Frame для GUI приложения Excel → JSON
class ExcelToJsonFrame(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)

        self.language = "ru"  # Язык интерфейса
        self.selected_files = []  # Выбранные Excel файлы
        self.save_folder = ""  # Папка для сохранения JSON
        self.searcher = None  # Экземпляр ExcelSearcher

        # Тексты интерфейса
        self.texts = {
            "ru": {
                "select_excel": "📂 Выбрать Excel",
                "delete_selected": "🗑 Удалить выбранное",
                "delete_all": "❌ Удалить все",
                "search": "🔍 Поиск",
                "save_json": "💾 Сохранить в JSON",
                "select_folder": "📁 Выбрать место для сохранения",
                "no_path": "Путь для сохранения не выбран",
                "lang_btn": "EN",
                "save_info": "Выбрано место сохранения:\n",
                "mode_label": "Выберите режим:",
                "modes": [
                    "🔍 Поиск по слову",
                    "🧱 Достать весь текст из столбцов",
                    "🆔 Достать весь текст из столбцов index",
                    "📏 Достать весь текст из строк",
                    "🆔 Достать весь текст из строк index",
                    "🔑 По двум столбцам",
                    "🆔 По двум столбцам index",
                    "📦 Сохранить все данные",
                ],
                "input_label": "Ключ:",
                "input_label_2": "Данные:",
                "msg_no_files": "Нет выбранных Excel файлов",
                "msg_enter_column": "Введите название столбца",
                "msg_enter_column_by_index": "Введите индекс столбца",
                "msg_found_count": "Найдено совпадений: {}",
                "msg_found_column": "Найдено записей в столбце: {}",
                "msg_found_rows": "Найдено совпадений: {}",
                "msg_found_rows_by_index": "Найдено строк по индексу: {}",
                "msg_found_all": "Всего записей: {}",
                "msg_saved": "Файл(ы) успешно сохранены",
                "msg_save_error": "Не выбрана папка для сохранения",
                "msg_save_info": "Сохранение возможно только для слова, столбца, строк, всего или двух столбцов"
            },
            "en": {
                "select_excel": "📂 Select Excel",
                "delete_selected": "🗑 Delete Selected",
                "delete_all": "❌ Delete All",
                "search": "🔍 Search",
                "save_json": "💾 Save to JSON",
                "select_folder": "📁 Select save folder",
                "no_path": "Save path not selected",
                "lang_btn": "RU",
                "save_info": "Selected save path:\n",
                "mode_label": "Select mode:",
                "modes": [
                    "🔍 Search by word",
                    "🧱 Extract all text from columns",
                    "🆔 Extract all text from columns index",
                    "📏 Extract all text from rows",
                    "🆔 Extract all text from rows index",
                    "🔑 By two columns",
                    "🆔 By two columns index",
                    "📦 Save all data",
                ],
                "input_label": "Key:",
                "input_label_2": "Data:",
                "msg_no_files": "No Excel files selected",
                "msg_enter_column": "Enter column name",
                "msg_enter_column_by_index": "Enter column index",
                "msg_found_count": "Matches found: {}",
                "msg_found_column": "Records found in column: {}",
                "msg_found_rows": "Matches found: {}",
                "msg_found_rows_by_index": "Rows found by index: {}",
                "msg_found_all": "Total records: {}",
                "msg_saved": "File(s) saved successfully",
                "msg_save_error": "Save folder not selected",
                "msg_save_info": "Saving is possible only for word, column, rows, all, or two columns"
            }
        }

        # 1. Кнопка выбора Excel
        self.select_excel_btn = Button(self, text=self.t("select_excel"), command=self.load_excel, bg="#87CEFA")
        self.select_excel_btn.pack(pady=(10, 5))

        # 2. Список выбранных файлов
        files_frame = tk.Frame(self)
        files_frame.pack(pady=(0, 10))
        self.file_listbox = Listbox(files_frame, width=60, height=4, selectmode=tk.SINGLE)
        self.file_listbox.grid(row=0, column=0, columnspan=2, padx=10)
        self.delete_selected_btn = Button(files_frame, text=self.t("delete_selected"), command=self.remove_selected, bg="#FFB6C1")
        self.delete_selected_btn.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.delete_all_btn = Button(files_frame, text=self.t("delete_all"), command=self.clear_all, bg="#FFB6C1")
        self.delete_all_btn.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 3. Выбор режима
        mode_frame = tk.Frame(self)
        mode_frame.pack(pady=(5, 10))
        self.mode_label = tk.Label(mode_frame, text=self.t("mode_label"), font=("Segoe UI", 10, "bold"))
        self.mode_label.pack()
        self.selected_mode = tk.StringVar(value=self.t("modes")[0])
        self.mode_menu = tk.OptionMenu(mode_frame, self.selected_mode, *self.t("modes"), command=self.toggle_second_input)
        self.mode_menu.config(width=40, font=("Segoe UI", 10))
        self.mode_menu.pack(pady=(3, 8))

        # 4. Поля ввода
        input_frame = tk.Frame(self)
        input_frame.pack(pady=(0, 15))
        self.input_label = tk.Label(input_frame, text=self.t("input_label"), font=("Segoe UI", 10, "bold"))
        self.input_label.grid(row=0, column=0, padx=5)
        self.input_var = tk.StringVar()
        self.input_entry = tk.Entry(input_frame, textvariable=self.input_var, width=20, font=("Segoe UI", 11))
        self.input_entry.grid(row=0, column=1, padx=5)
        self.input_label2 = tk.Label(input_frame, text=self.t("input_label_2"), font=("Segoe UI", 10, "bold"))
        self.input_var2 = tk.StringVar()
        self.input_entry2 = tk.Entry(input_frame, textvariable=self.input_var2, width=20, font=("Segoe UI", 11))

        # 5. Кнопки
        self.search_btn = Button(self, text=self.t("search"), command=self.search_action, bg="#90EE90")
        self.search_btn.pack(pady=(5, 10))

        self.select_folder_btn = Button(self, text=self.t("select_folder"), command=self.select_folder, bg="#FFD700")
        self.select_folder_btn.pack(pady=(0, 5))
        self.save_path_var = tk.StringVar(value=self.t("no_path"))
        self.save_label = tk.Label(self, textvariable=self.save_path_var, font=("Segoe UI", 9), fg="gray")
        self.save_label.pack(pady=(0, 15))

        self.save_btn = Button(self, text=self.t("save_json"), command=self.save_json, bg="#FFA500")
        self.save_btn.pack(pady=(0, 10))

        self.lang_btn = Button(self, text=self.t("lang_btn"), command=self.switch_language, bg="#D8BFD8")
        self.lang_btn.pack(pady=(0, 10))

    def toggle_second_input(self, mode=None):
        """
        Управляет отображением второго поля ввода
        — показывается только при режимах 'По двум столбцам' и 'По двум столбцам index'
        """
        if mode is None:
            mode = self.selected_mode.get()

        two_column_modes = [
            "🔑 По двум столбцам",
            "🆔 По двум столбцам index",
            "🔑 By two columns",
            "🆔 By two columns index"
        ]

        if mode in two_column_modes:
            # показываем оба поля
            self.input_label.grid(row=0, column=0, padx=5)
            self.input_entry.grid(row=0, column=1, padx=5)
            self.input_label2.grid(row=0, column=2, padx=5)
            self.input_entry2.grid(row=0, column=3, padx=5)
        else:
            # показываем только одно
            self.input_label.grid(row=0, column=0, padx=5)
            self.input_entry.grid(row=0, column=1, padx=5)
            self.input_label2.grid_forget()
            self.input_entry2.grid_forget()

    def t(self, key):
        return self.texts[self.language][key]

    def switch_language(self):
        self.language = "en" if self.language == "ru" else "ru"
        self.update_texts()

    def update_texts(self):
        self.select_excel_btn.config(text=self.t("select_excel"))
        self.delete_selected_btn.config(text=self.t("delete_selected"))
        self.delete_all_btn.config(text=self.t("delete_all"))
        self.search_btn.config(text=self.t("search"))
        self.save_btn.config(text=self.t("save_json"))
        self.select_folder_btn.config(text=self.t("select_folder"))
        self.lang_btn.config(text=self.t("lang_btn"))
        self.mode_label.config(text=self.t("mode_label"))
        self.input_label.config(text=self.t("input_label"))
        self.input_label2.config(text=self.t("input_label_2"))

        # обновляем выпадающее меню режимов
        menu = self.mode_menu["menu"]
        menu.delete(0, "end")
        for mode in self.t("modes"):
            menu.add_command(label=mode,
                             command=lambda m=mode: [self.selected_mode.set(m), self.toggle_second_input(m)])

        current_mode = self.selected_mode.get()
        if current_mode not in self.t("modes"):
            current_mode = self.t("modes")[0]
            self.selected_mode.set(current_mode)

        # ✅ обновляем отображение второго поля ввода
        self.toggle_second_input(current_mode)

        # ✅ обновляем текст под надписью пути
        if not self.save_folder:  # если путь ещё не выбран
            self.save_path_var.set(self.t("no_path"))
        else:
            # если путь выбран, показываем ту же надпись, но с новым переводом
            self.save_path_var.set(f"{self.t('save_info')}{self.save_folder}")

    def load_excel(self):
        files = filedialog.askopenfilenames(
            title=self.t("select_excel"),
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if files:
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_listbox.insert(END, f)
            self.searcher = ExcelSearcher(self.selected_files)

    def remove_selected(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            removed_file = self.file_listbox.get(index)
            self.file_listbox.delete(index)
            self.selected_files.remove(removed_file)
            self.searcher = ExcelSearcher(self.selected_files) if self.selected_files else None

    def clear_all(self):
        self.file_listbox.delete(0, END)
        self.selected_files.clear()
        self.searcher = None

    def select_folder(self):
        folder = filedialog.askdirectory(title=self.t("select_folder"))
        if folder:
            self.save_folder = folder
            self.save_path_var.set(f"{self.t('save_info')}{folder}")

    def search_action(self):
        if not self.searcher:
            messagebox.showwarning("Warning", self.t("msg_no_files"))
            return

        mode = self.selected_mode.get()
        query = self.input_var.get().strip()
        query2 = self.input_var2.get().strip()

        # 🔍 Поиск по слову
        if mode in ["🔍 Поиск по слову", "🔍 Search by word"]:
            results, total_found = self.searcher.search_word(query)
            messagebox.showinfo("Result", self.t("msg_found_count").format(total_found))

        # 🧱 Достать весь текст из столбцов
        elif mode in ["🧱 Достать весь текст из столбцов", "🧱 Extract all text from columns"]:
            if not query:
                messagebox.showwarning("Warning", self.t("msg_enter_column"))
                return
            results, total_found = self.searcher.search_column(query)
            messagebox.showinfo("Result", self.t("msg_found_column").format(total_found))

        # 🆔 Достать весь текст из столбцов index
        elif mode in ["🆔 Достать весь текст из столбцов index", "🆔 Extract all text from columns index"]:
            if not query:
                messagebox.showwarning("Warning", self.t("msg_enter_column_by_index"))
                return
            results, total_found = self.searcher.search_column_by_index(query)
            messagebox.showinfo("Result", self.t("msg_found_column").format(total_found))

        # 📏 Достать весь текст из строк
        elif mode in ["📏 Достать весь текст из строк", "📏 Extract all text from rows"]:
            results, total_found = self.searcher.search_rows(query)
            messagebox.showinfo("Result", self.t("msg_found_rows").format(total_found))

        # 🆔 Достать весь текст из строк index
        elif mode in ["🆔 Достать весь текст из строк index", "🆔 Extract all text from rows index"]:
            if not query.isdigit():
                messagebox.showwarning("Warning", "Введите числовой индекс строки")
                return
            results, total_found = self.searcher.search_rows_by_index(int(query))
            messagebox.showinfo("Result", self.t("msg_found_rows_by_index").format(total_found))

        # 🔑 По двум столбцам
        elif mode in ["🔑 По двум столбцам", "🔑 By two columns"]:
            if not query or not query2:
                messagebox.showwarning("Warning", self.t("msg_enter_column"))
                return
            results, total_found = self.searcher.search_two_columns(query, query2)
            messagebox.showinfo("Result", f"Найдено ключ-значений: {total_found}")

        # 🆔 По двум столбцам index
        elif mode in ["🆔 По двум столбцам index", "🆔 By two columns index"]:
            if not query or not query2:
                messagebox.showwarning("Warning", self.t("msg_enter_column_by_index"))
                return
            results, total_found = self.searcher.search_two_columns_by_index(query, query2)
            messagebox.showinfo("Result", f"Найдено ключ-значений: {total_found}")

        # 📦 Сохранить все данные
        elif mode in ["📦 Сохранить все данные", "📦 Save all data"]:
            results, total_found = self.searcher.get_all_data()
            messagebox.showinfo("Result", self.t("msg_found_all").format(total_found))

        else:
            messagebox.showinfo("Info", self.t("msg_save_info"))
            return

    def save_json(self):
        if not self.searcher:
            messagebox.showwarning("Warning", self.t("msg_no_files"))
            return

        if not self.save_folder:
            messagebox.showwarning("Warning", self.t("msg_save_error"))
            return

        mode = self.selected_mode.get()
        query = self.input_var.get().strip()
        query2 = self.input_var2.get().strip()
        data_to_save = {}

        # 🔍 Поиск по слову
        if mode in ["🔍 Поиск по слову", "🔍 Search by word"]:
            data_to_save = {query: self.searcher.search_word(query)}

        # 🧱 Поиск по названию столбца
        elif mode in ["🧱 Достать весь текст из столбцов", "🧱 Extract all text from columns"]:
            data_to_save = self.searcher.search_column(query)

        # 🆔 Поиск по индексу столбца
        elif mode in ["🆔 Достать весь текст из столбцов index", "🆔 Extract all text from columns index"]:
            data_to_save = self.searcher.search_column_by_index(query)

        # 📏 Поиск по строкам
        elif mode in ["📏 Достать весь текст из строк", "📏 Extract all text from rows"]:
            data_to_save = {"rows": self.searcher.search_rows(query)}

        # 🆔 Поиск строк по индексу
        elif mode in ["🆔 Достать весь текст из строк index", "🆔 Extract all text from rows index"]:
            if not query.isdigit():
                messagebox.showwarning("Warning", "Введите числовой индекс строки")
                return
            data_to_save = {"rows": self.searcher.search_rows_by_index(int(query))}

        # 🔑 По двум столбцам
        elif mode in ["🔑 По двум столбцам", "🔑 By two columns"]:
            data_to_save = self.searcher.search_two_columns(query, query2)

        # 🆔 По двум столбцам index
        elif mode in ["🆔 По двум столбцам index", "🆔 By two columns index"]:
            data_to_save = self.searcher.search_two_columns_by_index(query, query2)

        # 📦 Сохранить всё
        elif mode in ["📦 Сохранить все данные", "📦 Save all data"]:
            data_to_save = self.searcher.get_all_data()

        else:
            messagebox.showinfo("Info", self.t("msg_save_info"))
            return

        # 💾 Сохранение результата
        for file in self.selected_files:
            base_name = os.path.splitext(os.path.basename(file))[0]
            save_path = os.path.join(self.save_folder, f"{base_name}.json")
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)

        messagebox.showinfo("Saved", self.t("msg_saved"))

# Запуск функции
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel → JSON Converter")
    try:
        pil_image = Image.open('ico.png')
        icon = ImageTk.PhotoImage(pil_image)
        root.iconphoto(True, icon)
    except Exception as e:
        print(f"Не удалось установить иконку: {e}")
    frame = ExcelToJsonFrame(root)
    frame.pack(padx=10, pady=10)
    root.mainloop()