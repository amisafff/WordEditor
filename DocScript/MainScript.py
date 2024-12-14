import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

from DocScript import Begdans
from DocScript import tamplates

template_paths = [
    'DocFile/1-1. Бегданс _Протокол ГЭК_для КП.docx',
    'DocFile/График защиты ДП _на дверь_для КП.docx',
    'DocFile/График защиты ДП _с оценками_15.06.24_ для КП.docx',
    'DocFile/Присвоение квалификации_№1_15.06.24_для КП.docx',
    'DocFile/Списки студентов, тем, оценки рецензентов_для членов ГЭК_ для КП.docx',
    'DocFile/Статистика ГЭК_для КП.docx'
]


def mainscr(self, callback , output_folder):
    root = tk.Tk()
    root.withdraw()
    data1 = self.data
    print(data1)

    if output_folder:
        # Добавление сохранения в папку с текущей датой и временем
        current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        timestamped_folder = os.path.join(output_folder, current_datetime)
        os.makedirs(timestamped_folder, exist_ok=True)

        # Создание папки "студенты" внутри папки с текущей датой
        students_folder = os.path.join(timestamped_folder, "студенты")
        os.makedirs(students_folder, exist_ok=True)

        # Сохранение в папку "студенты"
        Begdans.fill_word_template(self.data, template_paths[0], students_folder)
        print(f"Файл сохранен в папку: {students_folder}")

        # Сохранение остальных файлов в папку с текущей датой
        tamplates.run_async(self.data, timestamped_folder)
        print(f"Выбрана папка для сохранения: {timestamped_folder}")
        callback(self, True)
    else:
        print("Сохранение отменено.")
        callback(self, False)
