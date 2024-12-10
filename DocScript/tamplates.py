import os
from docx import Document


def fill_word_template(dat_all, output_file_path):
    if not os.path.exists(output_file_path):
        os.makedirs(output_file_path)

    template_paths = [
        'DocFile/График защиты ДП _на дверь_для КП.docx',
        'DocFile/График защиты ДП _с оценками_15.06.24_ для КП.docx',
        'DocFile/Присвоение квалификации_№1_15.06.24_для КП.docx',
        'DocFile/Списки студентов, тем, оценки рецензентов_для членов ГЭК_ для КП.docx',
        'DocFile/Статистика ГЭК_для КП.docx'
    ]

    for template_index, template_path in enumerate(template_paths):

        doc = Document(template_path)

        data = {}

        for index, record in enumerate(dat_all):
            # Получаем строку с членами ГЭК
            geks_str = record[6]  # Столбец с членами ГЭК
            # Разделяем строку по запятой (или другим разделителем, если необходимо)
            geks = geks_str.split(',')  # Если члены разделены запятой
            # Убираем лишние пробелы
            geks = [geek.strip() for geek in geks]

            substringsDate = record[1].split('-')
            # Сохраняем каждого члена ГЭК
            for g_index, member in enumerate(geks):
                data[f"Члены_ГЭК{index + 1}_{g_index + 1}"] = member
            data[f"ФИО{index + 1}"] = record[2]  # Столбец "ФИО"
            data[f"Специальность"] = record[3]  # Столбец "Специальность"
            data[f"Тема_ДП{index + 1}"] = record[4]  # Столбец "Тема ДП"
            data[f"Председатель_ГЭК"] = record[5]  # Столбец "Председатель ГЭК"
            data[f"Консультант{index + 1}"] = record[7]  # Столбец "Консультант"
            data[f"ФОРМ{index + 1}"] = record[8]  # Столбец "Форма обучения"
            data[f"Руководитель{index + 1}"] = record[9]  # Столбец "Руководитель"
            data[f"ОЦ{index + 1}"] = record[10]  # Столбец "Оценка"
            data[f"ВИЗА{index + 1}"] = record[11]  # Столбец "Виза лица, составившего протокол"
            data[f"Степень{index + 1}"] = record[12]  # Столбец "Степень"
            data[f"ОТЛ{index + 1}"] = record[13]  # Столбец "Диплом с отличием"
            data[f"НОМ{index + 1}"] = record[0]  # Столбец "Номер протокола"
            data[f"День"] = substringsDate[2]
            data[f"Год"] = substringsDate[0]
            data[f"Время"] = substringsDate[3]

        for paragraph in doc.paragraphs:
            for key, value in data.items():
                if f"{{{{{key}}}}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", str(value))
                # Удаляем метку, если значения нет в data
                if f"{{{{{key}}}}}" in paragraph.text and key not in data:
                    paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", "")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in data.items():
                        if f"{{{{{key}}}}}" in cell.text:
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))
                        # Удаляем метку, если значения нет в data
                        if f"{{{{{key}}}}}" in cell.text and key not in data:
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", "")

        file_name = os.path.basename(template_path)
        save_path = os.path.join(output_file_path, file_name)
        doc.save(save_path)
        print(f"Файл {file_name} успешно сохранен в {output_file_path}.")
