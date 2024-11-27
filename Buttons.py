import tkinter as tk
from tkinter import ttk, simpledialog


def add_row(self):

    def save_dialog():
        data = []

        for i, entry in enumerate(entries):
            if i in [3, 6, 8]:  # Столбцы 4, 5, 6 (с 0-индекса) — выпадающие списки
                selected = [self.check_options[i + 1][j] for j, var in enumerate(vars[i]) if var.get()]
                data.append(", ".join(selected) if selected else "")
            elif i in [12, 13]:  # Столбцы 13, 14 — бинарный выбор
                data.append(entry.get() if entry.get() else "Нет")
            else:
                data.append(entry.get() if entry.get() else "")

        # Добавить строку в таблицу и структуру данных

        self.tree.insert("", tk.END, values=data)
        self.data.append(data)

        dialog.destroy()

    dialog = tk.Toplevel(self)
    dialog.title("Добавить строку")
    dialog.geometry("600x500")
    dialog.configure(bg="#f8f9fa")

    entries = []
    vars = {}
    columns = [
        "Номер протокола", "Месяц год время", "ФИО", "Специальность", "Тема ДП",
        "Председатель ГЭК", "Члены ГЭК", "Консультант", "Форма обучения", "Руководитель",
        "Оценка", "Виза лица, составившего протокол ", "Степень", "Диплом с отличием"
    ]
    for i in range(14):  # Всего 14 столбцов
        tk.Label(dialog, text=f"{columns[i]}:", bg="#f8f9fa").grid(row=i, column=0, pady=5, padx=10, sticky="w")
        if i in [3, 6, 8]:  # Столбцы 4, 5, 6 — выпадающие списки
            frame = tk.Frame(dialog, bg="#f8f9fa")
            frame.grid(row=i, column=1, pady=5, padx=10, sticky="w")
            vars[i] = []
            for option in self.check_options[i + 1]:
                var = tk.BooleanVar()
                tk.Checkbutton(frame, text=option, variable=var, bg="#f8f9fa").pack(side=tk.LEFT, padx=5)
                vars[i].append(var)
            entries.append(None)
        elif i in [12, 13]:  # Столбцы 13, 14 — бинарный выбор
            entry = ttk.Combobox(dialog, values=["Да", "Нет"])
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


def add_option(self):
    column_index = simpledialog.askinteger("Добавить опцию",
                                           "Для какого столбца (4, 5 или 6) добавить новую опцию?")
    if column_index in self.check_options:
        new_option = simpledialog.askstring("Добавить опцию", "Введите новую опцию:")
        if new_option:
            self.check_options[column_index].append(new_option)


