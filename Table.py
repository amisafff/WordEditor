import tkinter as tk
from tkinter import ttk
import Buttons
from DocScript import MainScript


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Word Editor")
        self.geometry("1920x900")
        self.configure(bg="#f8f9fa")

        # Структура для хранения данных
        self.data = []

        # Создание навигационного бара
        navbar = tk.Frame(self, bg="#007bff")
        navbar.pack(side=tk.TOP, fill=tk.X)

        title_label = tk.Label(navbar, text="Таблица данных", bg="#007bff", fg="white", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # Словарь для хранения значений чекбоксов для выпадающих списков
        self.check_options = {4: ["Опция 1", "Опция 2", "Опция 3"],
                              7: ["Опция A", "Опция B", "Опция C"],
                              9: ["Опция X", "Опция Y", "Опция Z"]}

        columns = [
            "Номер протокола", "Месяц год время", "ФИО", "Специальность", "Тема ДП",
            "Председатель ГЭК", "Члены ГЭК", "Консультант","Форма обучения","Руководитель",
            "Оценка", "Виза лица, составившего протокол ", "Степень", "Диплом с отличием"
        ]

        # Создание таблицы с 14 столбцами
        self.tree = ttk.Treeview(self, columns=[f"Col {i}" for i in range(1, 15)], show='headings', height=15)
        for i in range(1, 15):
            self.tree.heading(f"Col {i}", text=columns[i - 1])
            self.tree.column(f"Col {i}", width=90, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=(20, 0))

        # Кнопки управления
        button_frame = tk.Frame(self, bg="#f8f9fa")
        button_frame.pack(pady=10)

        add_button = tk.Button(button_frame, text="Добавить строку", command=lambda: Buttons.add_row(self),
                               bg="#28a745",
                               fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        add_button.grid(row=0, column=0, padx=5)

        delete_button = tk.Button(button_frame, text="Удалить строку", command=lambda: Buttons.delete_row(self),
                                  bg="#dc3545",
                                  fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        delete_button.grid(row=0, column=1, padx=5)

        add_option_button = tk.Button(button_frame, text="Добавить опцию", command=lambda: Buttons.add_option(self),
                                      bg="#007bff",
                                      fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        add_option_button.grid(row=0, column=2, padx=5)

        edit_button = tk.Button(button_frame, text="Изменить строку", command=lambda: Buttons.edit_row(self),
                                bg="#ffc107",
                                fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        edit_button.grid(row=0, column=3, padx=5)

        save_button = tk.Button(button_frame, text="Сохранить в Word", command=lambda: MainScript.mainscr(self),
                                bg="#87CEFA",
                                fg="white", font=("Arial", 9, "bold"), padx=5, pady=5)
        save_button.grid(row=0, column=4, padx=5)



