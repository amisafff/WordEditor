import tkinter as tk
from tkinter import ttk, simpledialog

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word Editor")
        self.geometry("1000x500")
        self.configure(bg="#f8f9fa")

        # Создание навигационного бара
        navbar = tk.Frame(self, bg="#007bff")
        navbar.pack(side=tk.TOP, fill=tk.X)

        title_label = tk.Label(navbar, text="Таблица данных", bg="#007bff", fg="white", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # Словарь для хранения значений чекбоксов для выпадающих списков
        self.check_options = {4: ["Опция 1", "Опция 2", "Опция 3"],
                              5: ["Опция A", "Опция B", "Опция C"],
                              6: ["Опция X", "Опция Y", "Опция Z"]}

        # Создание таблицы с 10 столбцами
        self.tree = ttk.Treeview(self, columns=[f"Col {i}" for i in range(1, 11)], show='headings', height=15)
        for i in range(1, 11):
            self.tree.heading(f"Col {i}", text=f"Столбец {i}")
            self.tree.column(f"Col {i}", width=90, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=(20, 0))

        # Кнопки управления
        button_frame = tk.Frame(self, bg="#f8f9fa")
        button_frame.pack(pady=10)

        add_button = tk.Button(button_frame, text="Добавить строку", command=self.add_row, bg="#28a745", fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        add_button.grid(row=0, column=0, padx=5)

        delete_button = tk.Button(button_frame, text="Удалить строку", command=self.delete_row, bg="#dc3545", fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        delete_button.grid(row=0, column=1, padx=5)

        add_option_button = tk.Button(button_frame, text="Добавить опцию", command=self.add_option, bg="#007bff", fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        add_option_button.grid(row=0, column=2, padx=5)

        # Кнопка для редактирования строки
        edit_button = tk.Button(button_frame, text="Изменить строку", command=self.edit_row, bg="#ffc107", fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        edit_button.grid(row=0, column=3, padx=5)

    def add_row(self):
        # Функция для добавления строки через диалоговое окно
        def save_dialog():
            data = []
            for i, entry in enumerate(entries):
                if i in [3, 4, 5]:  # Столбцы 4, 5, 6 (с 0-индекса) — выпадающие списки
                    selected = [self.check_options[i + 1][j] for j, var in enumerate(vars[i]) if var.get()]
                    data.append(", ".join(selected) if selected else "")  # Если ничего не выбрано, добавляем пустую строку
                else:  # Для обычных текстовых полей
                    data.append(entry.get() if entry.get() else "")  # Добавляем пустую строку, если поле пустое
            self.tree.insert("", tk.END, values=data)
            dialog.destroy()

        dialog = tk.Toplevel(self)
        dialog.title("Добавить строку")
        dialog.geometry("600x400")
        dialog.configure(bg="#f8f9fa")

        entries = []
        vars = {}

        for i in range(10):
            tk.Label(dialog, text=f"Столбец {i + 1}:", bg="#f8f9fa").grid(row=i, column=0, pady=5, padx=10, sticky="w")
            if i in [3, 4, 5]:  # Столбцы 4, 5, 6 — выпадающие списки
                frame = tk.Frame(dialog, bg="#f8f9fa")
                frame.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                vars[i] = []
                for option in self.check_options[i + 1]:
                    var = tk.BooleanVar()
                    tk.Checkbutton(frame, text=option, variable=var, bg="#f8f9fa").pack(side=tk.LEFT, padx=5)
                    vars[i].append(var)
                entries.append(None)  # Записываем None, потому что для этих столбцов используем чекбоксы, а не текстовое поле
            else:  # Для обычных текстовых полей
                entry = tk.Entry(dialog, width=30)
                entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                entries.append(entry)

        tk.Button(dialog, text="Сохранить", command=save_dialog, bg="#28a745", fg="white").grid(row=10, column=0, columnspan=2, pady=20)

    def edit_row(self):
        # Функция для редактирования выбранной строки
        selected_item = self.tree.selection()
        if selected_item:
            current_values = self.tree.item(selected_item, "values")

            # Открытие диалога для редактирования
            def save_dialog():
                data = []
                for i, entry in enumerate(entries):
                    if i in [3, 4, 5]:  # Столбцы 4, 5, 6 (с 0-индекса) — выпадающие списки
                        selected = [self.check_options[i + 1][j] for j, var in enumerate(vars[i]) if var.get()]
                        data.append(", ".join(selected) if selected else "")
                    else:  # Для обычных текстовых полей
                        data.append(entry.get() if entry.get() else "")
                self.tree.item(selected_item, values=data)
                dialog.destroy()

            dialog = tk.Toplevel(self)
            dialog.title("Изменить строку")
            dialog.geometry("600x400")
            dialog.configure(bg="#f8f9fa")

            entries = []
            vars = {}

            for i in range(10):
                tk.Label(dialog, text=f"Столбец {i + 1}:", bg="#f8f9fa").grid(row=i, column=0, pady=5, padx=10, sticky="w")
                if i in [3, 4, 5]:  # Столбцы 4, 5, 6 — выпадающие списки
                    frame = tk.Frame(dialog, bg="#f8f9fa")
                    frame.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                    vars[i] = []
                    selected_values = current_values[i].split(", ")
                    for option in self.check_options[i + 1]:
                        var = tk.BooleanVar(value=(option in selected_values))
                        tk.Checkbutton(frame, text=option, variable=var, bg="#f8f9fa").pack(side=tk.LEFT, padx=5)
                        vars[i].append(var)
                    entries.append(None)
                else:  # Для обычных текстовых полей
                    entry = tk.Entry(dialog, width=30)
                    entry.grid(row=i, column=1, pady=5, padx=10, sticky="w")
                    entry.insert(0, current_values[i])
                    entries.append(entry)

            tk.Button(dialog, text="Сохранить", command=save_dialog, bg="#28a745", fg="white").grid(row=10, column=0, columnspan=2, pady=20)
        else:
            simpledialog.messagebox.showinfo("Ошибка", "Выберите строку для редактирования")

    def delete_row(self):
        # Удаление выбранной строки
        selected_item = self.tree.selection()
        if selected_item:
            self.tree.delete(selected_item)

    def add_option(self):
        # Диалог для ввода новой опции
        column = simpledialog.askinteger("Добавить опцию", "Выберите столбец (4, 5 или 6)", minvalue=4, maxvalue=6)
        if column and column in [4, 5, 6]:
            option = simpledialog.askstring("Новая опция", "Введите название опции")
            if option:
                self.check_options[column].append(option)

if __name__ == "__main__":
    app = App()
    app.mainloop()
