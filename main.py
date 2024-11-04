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
        self.check_options = {
            4: ["Опция 1", "Опция 2", "Опция 3"],
            5: ["Опция A", "Опция B", "Опция C"],
            6: ["Опция X", "Опция Y", "Опция Z"]
        }

        # Создание таблицы с 10 столбцами
        self.tree = ttk.Treeview(self, columns=[f"Col {i}" for i in range(1, 11)], show='headings', height=15)
        for i in range(1, 11):
            self.tree.heading(f"Col {i}", text=f"Столбец {i}")
            self.tree.column(f"Col {i}", width=90, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=(20, 0))

        # Кнопки управления
        button_frame = tk.Frame(self, bg="#f8f9fa")
        button_frame.pack(pady=10)

        add_button = tk.Button(button_frame, text="Добавить строку", command=self.add_row, bg="#28a745", fg="white",
                               font=("Arial", 9, "bold"), padx=3, pady=3)
        add_button.grid(row=0, column=0, padx=5)

        delete_button = tk.Button(button_frame, text="Удалить строку", command=self.delete_row, bg="#dc3545",
                                  fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        delete_button.grid(row=0, column=1, padx=5)

        add_option_button = tk.Button(button_frame, text="Добавить опцию", command=self.add_option, bg="#007bff",
                                      fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        add_option_button.grid(row=0, column=2, padx=5)

        # Событие для редактирования ячеек
        self.tree.bind("<Double-1>", self.on_double_click)

    def add_row(self):
        # Создаем новую строку в таблице
        self.tree.insert("", tk.END, values=["" for _ in range(10)])

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

    def on_double_click(self, event):
        # Получаем строку и столбец, по которым был двойной щелчок
        item_id = self.tree.focus()
        column = self.tree.identify_column(event.x)
        col_index = int(column.replace("#", ""))

        # Если столбец имеет ограниченный набор опций, используем Combobox
        if col_index in self.check_options:
            self.create_combobox(item_id, col_index)
        else:
            self.create_entry(item_id, col_index)

    def create_entry(self, item_id, col_index):
        # Создаем Entry для ввода текста
        x, y, width, height = self.tree.bbox(item_id, f"Col {col_index}")
        entry = tk.Entry(self.tree, font=("Arial", 10))
        entry.place(x=x, y=y, width=width, height=height)

        # Устанавливаем текущее значение
        entry.insert(0, self.tree.item(item_id, "values")[col_index - 1])

        # Фокус и сохранение значения по завершению
        entry.focus()
        entry.bind("<Return>", lambda e: self.save_entry_value(item_id, col_index, entry))

    def save_entry_value(self, item_id, col_index, entry):
        # Сохранение значения в ячейку
        self.tree.set(item_id, f"Col {col_index}", entry.get())
        entry.destroy()  # Удаляем Entry

    def create_combobox(self, item_id, col_index):
        # Создаем Combobox для выбора из ограниченного набора опций
        x, y, width, height = self.tree.bbox(item_id, f"Col {col_index}")
        combobox = ttk.Combobox(self.tree, values=self.check_options[col_index], font=("Arial", 10))
        combobox.place(x=x, y=y, width=width, height=height)

        # Устанавливаем текущее значение
        combobox.set(self.tree.item(item_id, "values")[col_index - 1])

        # Сохранение выбранного значения по завершению
        combobox.bind("<<ComboboxSelected>>", lambda e: self.save_combobox_value(item_id, col_index, combobox))
        combobox.focus()

    def save_combobox_value(self, item_id, col_index, combobox):
        # Сохранение значения из Combobox в ячейку
        self.tree.set(item_id, f"Col {col_index}", combobox.get())
        combobox.destroy()  #


if __name__ == "__main__":
    app = App()
    app.mainloop()