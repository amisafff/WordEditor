import tkinter as tk
from tkinter import ttk, simpledialog
from datetime import datetime


def add_row(self):
    def save_dialog():


        data = []

        for i, entry in enumerate(entries):
            if i == 1:  # Для столбца "Месяц год время"
                date_str = f"{year_combobox.get()}-{month_combobox.get()}-{day_combobox.get()}-{hour_combobox.get()}:{minute_combobox.get()}"
                data.append(date_str)
                # Сохранить стандартное значение даты и времени
                self.default_values["date"] = (day_combobox.get(), month_combobox.get(), year_combobox.get())
                self.default_values["time"] = (hour_combobox.get(), minute_combobox.get())
                # Сохранить стандартное значение
                self.default_values["date"] = (day_combobox.get(), month_combobox.get(), year_combobox.get())
            elif i in [3, 6, 8]:  # Столбцы 4, 5, 6 — выпадающие списки
                selected = [self.check_options[i + 1][j] for j, var in enumerate(vars[i]) if var.get()]
                data.append(", ".join(selected) if selected else "")
                self.default_values[f"options_{i}"] = [var.get() for var in vars[i]]
                # Сохранить стандартное значение
                self.default_values[f"options_{i}"] = [var.get() for var in vars[i]]
            elif i in [5, 11]:  # Поля "Председатель ГЭК" и "Виза лица, составившего протокол"
                value = entry.get() if entry.get() else ""
                data.append(value)
                # Сохранить стандартное значение
                self.default_values[f"text_{i}"] = value
            elif i in [12, 13]:  # Столбцы 13, 14 — бинарный выбор
                value = entry.get() if entry.get() else "Нет"
                data.append(value)
                # Сохранить стандартное значение
                self.default_values[f"binary_{i}"] = value
            elif i == 6:  # Поле "Члены ГЭК" (многострочная организация)
                selected = [self.check_options[i + 1][j] for j, var in enumerate(vars[i]) if var.get()]
                data.append(", ".join(selected) if selected else "")
                self.default_values[f"options_{i}"] = [var.get() for var in vars[i]]
            else:
                value = entry.get() if entry.get() else ""
                data.append(value)

        # Добавить строку в таблицу и структуру данных
        self.tree.insert("", tk.END, values=data)
        self.data.append(data)

        dialog.destroy()

    dialog = tk.Toplevel(self)
    dialog.title("Добавить строку")
    dialog.geometry("600x500")
    dialog.configure(bg="#f8f9fa")

    # Инициализация стандартных значений, если их нет
    if not hasattr(self, "default_values"):
        self.default_values = {}



    entries = []
    vars = {}
    columns = [
        "Номер протокола", "Месяц год время", "ФИО", "Специальность", "Тема ДП",
        "Председатель ГЭК", "Члены ГЭК", "Консультант", "Форма обучения", "Руководитель",
        "Оценка", "Виза лица, составившего протокол", "Степень", "Диплом с отличием"
    ]
    for i in range(14):

        tk.Label(dialog, text=f"{columns[i]}:", bg="#f8f9fa").grid(row=i, column=0, pady=5, padx=10, sticky="w")

        if i == 1:  # Для столбца "Месяц год время"
            date_time_frame = tk.Frame(dialog, bg="#f8f9fa")
            date_time_frame.grid(row=i, column=1, pady=5, padx=10, sticky="w")

            # Установка стандартных значений даты и времени
            day = self.default_values.get("date", (str(datetime.now().day).zfill(2),))[0]
            month = self.default_values.get("date", ("", str(datetime.now().month).zfill(2)))[1]
            year = self.default_values.get("date", ("", "", str(datetime.now().year)))[2]
            hour = self.default_values.get("time", (str(datetime.now().hour).zfill(2),))[0]
            minute = self.default_values.get("time", (str(datetime.now().minute).zfill(2),))[0]

            # Выпадающие списки для дня, месяца, года
            day_combobox = ttk.Combobox(date_time_frame, values=[str(d).zfill(2) for d in range(1, 32)], width=5)
            day_combobox.set(day)
            day_combobox.pack(side=tk.LEFT, padx=5)

            month_combobox = ttk.Combobox(date_time_frame, values=[str(m).zfill(2) for m in range(1, 13)], width=5)
            month_combobox.set(month)
            month_combobox.pack(side=tk.LEFT, padx=5)

            year_combobox = ttk.Combobox(date_time_frame, values=[str(y) for y in range(2000, 2031)], width=7)
            year_combobox.set(year)
            year_combobox.pack(side=tk.LEFT, padx=5)

            # Выпадающие списки для часов и минут
            hour_combobox = ttk.Combobox(date_time_frame, values=[str(h).zfill(2) for h in range(24)], width=5)
            hour_combobox.set(hour)
            hour_combobox.pack(side=tk.LEFT, padx=5)

            minute_combobox = ttk.Combobox(date_time_frame, values=[str(m).zfill(2) for m in range(60)], width=5)
            minute_combobox.set(minute)
            minute_combobox.pack(side=tk.LEFT, padx=5)

            # Сохранение комбобоксов в список entries
            entries.append((day_combobox, month_combobox, year_combobox, hour_combobox, minute_combobox))


        elif i in [3, 6, 8]:  # Столбцы 4, 5, 6 — выпадающие списки
            frame = tk.Frame(dialog, bg="#f8f9fa")
            frame.grid(row=i, column=1, pady=5, padx=10, sticky="w")

            vars[i] = []
            default_values = self.default_values.get(f"options_{i}", [False] * len(self.check_options[i + 1]))

            for index, (option, default_value) in enumerate(zip(self.check_options[i + 1], default_values)):
                var = tk.BooleanVar(value=default_value)
                cb = tk.Checkbutton(frame, text=option, variable=var, bg="#f8f9fa")
                cb.grid(row=index // 3, column=index % 3, padx=5, pady=2, sticky="w")
                vars[i].append(var)

            entries.append(None)

        elif i in [5, 11]:  # Поля "Председатель ГЭК" и "Виза лица, составившего протокол"
            default_value = self.default_values.get(f"text_{i}", "")
            entry = tk.Entry(dialog, width=30)
            entry.insert(0, default_value)
            entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
            entries.append(entry)

        elif i in [12, 13]:  # Столбцы 13, 14 — бинарный выбор
            default_value = self.default_values.get(f"binary_{i}", "Нет")
            entry = ttk.Combobox(dialog, values=["Да", "Нет"])
            entry.set(default_value)
            entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
            entries.append(entry)

        else:  # Текстовые поля
            entry = tk.Entry(dialog, width=30)
            entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
            entries.append(entry)

    tk.Button(dialog, text="Сохранить", command=save_dialog, bg="#28a745", fg="white").grid(row=14, column=0,
                                                                                            columnspan=2, pady=20)



def edit_row(self):
    selected_item = self.tree.selection()
    if selected_item:
        current_values = self.tree.item(selected_item, "values")

        def save_dialog():
            data = []
            for i, entry in enumerate(entries):
                if i in [3, 6, 8]:
                    selected = [self.check_options[i + 1][j] for j, var in enumerate(vars[i]) if var.get()]
                    data.append(", ".join(selected) if selected else "")
                elif i in [12, 13]:
                    data.append(entry.get() if entry.get() else "Нет")
                else:
                    data.append(entry.get() if entry.get() else "")

            # Обновить строку в таблице и структуру данных
            self.tree.item(selected_item, values=data)
            row_index = self.tree.index(selected_item)
            self.data[row_index] = data
            dialog.destroy()

        dialog = tk.Toplevel(self)
        dialog.title("Изменить строку")
        dialog.geometry("600x500")
        dialog.configure(bg="#f8f9fa")

        entries = []
        vars = {}
        columns = [
            "Номер протокола", "Месяц год время", "ФИО", "Специальность", "Тема ДП",
            "Председатель ГЭК", "Члены ГЭК", "Консультант", "Форма обучения", "Руководитель",
            "Оценка", "Виза лица, составившего протокол ", "Степень", "Диплом с отличием"
        ]
        for i in range(14):
            tk.Label(dialog, text=f"{columns[i]}:", bg="#f8f9fa").grid(row=i, column=0, pady=5, padx=10, sticky="w")
            if i in [3, 6, 8]:
                frame = tk.Frame(dialog, bg="#f8f9fa")
                frame.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                vars[i] = []
                selected_values = current_values[i].split(", ")
                for option in self.check_options[i + 1]:
                    var = tk.BooleanVar(value=(option in selected_values))
                    tk.Checkbutton(frame, text=option, variable=var, bg="#f8f9fa").pack(side=tk.LEFT, padx=5)
                    vars[i].append(var)
                entries.append(None)
            elif i in [12, 13]:
                entry = ttk.Combobox(dialog, values=["Да", "Нет"])
                entry.set(current_values[i])
                entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                entries.append(entry)
            else:
                entry = tk.Entry(dialog, width=30)
                entry.insert(0, current_values[i])
                entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                entries.append(entry)

        tk.Button(dialog, text="Сохранить", command=save_dialog, bg="#28a745", fg="white").grid(row=14, column=0,
                                                                                                columnspan=2, pady=20)


def delete_row(self):
    selected_item = self.tree.selection()
    if selected_item:
        row_index = self.tree.index(selected_item)
        self.tree.delete(selected_item)
        del self.data[row_index]


def manage_options(self):
    def update_options_list():
        option_list.delete(0, tk.END)
        for option in self.check_options[column_index]:
            option_list.insert(tk.END, option)

    def add_new_option():
        new_option = new_option_entry.get().strip()
        if new_option and new_option not in self.check_options[column_index]:
            self.check_options[column_index].append(new_option)
            update_options_list()
            new_option_entry.delete(0, tk.END)
            save_options_to_file()  # Сохраняем опции после добавления

    def delete_selected_option():
        selected_items = option_list.curselection()
        if selected_items:
            for index in selected_items[::-1]:
                del self.check_options[column_index][index]
            update_options_list()
            save_options_to_file()  # Сохраняем опции после удаления

    def save_options_to_file():
        file_path = f"options_{column_index}.txt"  # Для каждого столбца свой файл
        try:
            with open(file_path, "w", encoding="utf-8") as file:
                for option in self.check_options[column_index]:
                    file.write(option + "\n")
        except Exception as e:
            tk.messagebox.showerror("Ошибка", f"Не удалось сохранить опции: {e}")



    column_index = simpledialog.askinteger(
        "Управление опциями", "Для какого столбца (4, 7 или 9) вы хотите управлять опциями?"
    )
    if column_index in self.check_options:
        dialog = tk.Toplevel(self)
        dialog.title(f"Управление опциями для столбца {column_index}")
        dialog.geometry("400x400")
        dialog.configure(bg="#f8f9fa")

        tk.Label(dialog, text=f"Опции для столбца {column_index}:", bg="#f8f9fa", font=("Arial", 10, "bold")).pack(pady=10)

        option_list_frame = tk.Frame(dialog, bg="#f8f9fa")
        option_list_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        option_list = tk.Listbox(option_list_frame, selectmode=tk.MULTIPLE, height=10, width=40)
        option_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        update_options_list()

        scrollbar = tk.Scrollbar(option_list_frame, orient=tk.VERTICAL, command=option_list.yview)
        option_list.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        new_option_label = tk.Label(dialog, text="Добавить новую опцию:", bg="#f8f9fa")
        new_option_label.pack(pady=(10, 0))
        new_option_entry = tk.Entry(dialog, width=30)
        new_option_entry.pack(pady=(0, 10))

        button_frame = tk.Frame(dialog, bg="#f8f9fa")
        button_frame.pack(pady=10)

        add_button = tk.Button(
            button_frame, text="Добавить", command=add_new_option, bg="#28a745", fg="white", font=("Arial", 9, "bold")
        )
        add_button.grid(row=0, column=0, padx=5)

        delete_button = tk.Button(
            button_frame, text="Удалить выбранные", command=delete_selected_option, bg="#dc3545", fg="white", font=("Arial", 9, "bold")
        )
        delete_button.grid(row=0, column=1, padx=5)



        close_button = tk.Button(
            dialog, text="Закрыть", command=dialog.destroy, bg="#007bff", fg="white", font=("Arial", 9, "bold")
        )
        close_button.pack(pady=10)
