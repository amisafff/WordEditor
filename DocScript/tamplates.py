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
            data[f"ФИО{index+1}"] = record[2]
            data[f"Специальность"] = record[3]
            data[f"Месяц_год_время"] = record[1]

        for paragraph in doc.paragraphs:
            for key, value in data.items():
                if f"{{{{{key}}}}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", str(value))

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in data.items():
                        if f"{{{{{key}}}}}" in cell.text:
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", str(value))

        file_name = f"{template_index + 1}-output.docx"
        save_path = os.path.join(output_file_path, file_name)
        doc.save(save_path)
        print(f"Файл {file_name} успешно сохранен в {output_file_path}.")



