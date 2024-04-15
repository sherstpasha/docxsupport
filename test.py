from docxtpl import DocxTemplate
import os


replacements_once = {
    "theme": "Программная система классификации жалоб поступающих в службу 005 с использованием векторной модели представления знаний и машинного обучения",
}

replacements_autors = {
    "Шерстнев Павел Александрович": {
        "subject_name": "Шерстнев Павел Александрович",
        "subject_address": "РФ, 662155, г. Ачинск, ул. Уличная, д. 312, кв. 32",
        "passport_series": "1234",
        "passport_number": "567890",
        "passport_issued_by": "ОУФМС России в г. Москве",
        "passport_issue_date": "01.01.2020",
        "passport_division_code": "123-456",
        "birthday": "11",
        "birthmonth": "01",
        "birthyear": "1999",
        "country": "Россия",
    },
    "Кожин Константин Дмитриевич": {
        "subject_name": "Кожин Константин Дмитриевич",
        "subject_address": "РФ, 662155, г. Красноярск, ул. Ленина, д. 22, кв. 33",
        "passport_series": "4321",
        "passport_number": "098765",
        "passport_issued_by": "ОУФМС России в г. Москве",
        "passport_issue_date": "02.02.2010",
        "passport_division_code": "321-654",
        "birthday": "13",
        "birthmonth": "03",
        "birthyear": "2002",
        "country": "Россия",
    },
}

file_paths = [
    "templates\Согласие_на_обработку_Шаблон.docx",
    "templates\Согласие_на_указание_Шаблон.docx",
]

save_path = ""

for file_path in file_paths:

    for author, replacements in replacements_autors.items():
        doc = DocxTemplate(file_path)
        doc.render(replacements | replacements_once)

        file_name = os.path.basename(file_path.replace("Шаблон", author))

        full_file_path = os.path.join(save_path, file_name)

        doc.save(full_file_path)
