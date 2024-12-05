import os
import tkinter as tk
from tkinter import filedialog

import Buttons
import Table
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


def mainscr(self):
    root = tk.Tk()
    root.withdraw()
    data1 = self.data
    print(data1)
    output_folder = filedialog.askdirectory(title="Выберите папку для сохранения документов")

    if output_folder:

        students_folder = os.path.join(output_folder, "студенты")
        os.makedirs(students_folder, exist_ok=True)

        Begdans.fill_word_template(self.data, template_paths[0], students_folder)
        print(f"Файл сохранен в папку: {students_folder}")

        tamplates.fill_word_template(self.data, output_folder)

        print(f"Выбрана папка для сохранения: {output_folder}")
    else:
        print("Сохранение отменено.")
