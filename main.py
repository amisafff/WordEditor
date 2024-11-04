import tkinter as tk
from tkinter import ttk, simpledialog

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Таблица с вводом данных и выпадающими списками")
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

    def add_row(self):
        # Создаем новую строку в таблице
        row_id = self.tree.insert("", tk.END, values=["" for _ in range(10)])

        # Массив для хранения Entry виджетов
        entries = []

        # Добавляем поля ввода для каждой ячейки новой строки
        for col in range(1, 11):
            if col in [4, 5, 6]:  # Выпадающий список с чекбоксами для 4, 5 и 6 столбцов
                button = tk.Menubutton(self, text="Выберите", relief=tk.RAISED, bg="#e9ecef", font=("Arial", 9))
                menu = tk.Menu(button, tearoff=0)
                button["menu"] = menu

                # Добавляем опции чекбоксов
                for option in self.check_options[col]:
                    var = tk.BooleanVar()
                    menu.add_checkbutton(label=option, variable=var)

                # Расположение выпадающего списка
                button_window = self.tree.bbox(row_id, f"Col {col}")
                button.place(x=button_window[0] + 20, y=button_window[1] + 30)
            else:  # Для остальных столбцов добавляем Entry
                entry = tk.Entry(self, font=("Arial", 10))
                entry_window = self.tree.bbox(row_id, f"Col {col}")
                entry.place(x=entry_window[0] + 20, y=entry_window[1] + 30)
                entries.append((col, entry))

        # Кнопка для сохранения введенной строки
        save_button = tk.Button(self, text="Сохранить", command=lambda: self.save_row(row_id, entries), bg="#28a745", fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        save_button.place(x=20, y=450)

    def save_row(self, row_id, entries):
        # Сохранение данных из Entry виджетов в таблицу
        data = [""] * 10
        for col, entry in entries:
            data[col - 1] = entry.get()  # Сохраняем значение в нужной ячейке строки
            entry.destroy()  # Убираем Entry после ввода данных
        self.tree.item(row_id, values=data)

    def delete_row(self):
        # Удаление выбранной строки
        selected_item = self.tree.selection()
        if selected_item:
            self.tree.delete(selected_item)

    def add_option(self):
        # Диалог для ввода новой опции
        column = simpledialog.askinteger("Добавить опцию", "Выберите столбец (4, 5 или 6)", minvalue=4, maxvalue=6)
        if column and column in [4, 5, 6]:
            option = simpledialog.askstring("Новая опция", "Введите название опциии")
            if option:
                self.check_options[column].append(option)

if __name__ == "__main__":
    app = App()
    app.mainloop()
