from docx import Document

from utils import append_text_to_docx, add_passport_and_sing



append_dict = {
    "Название программы для ЭВМ или базы данных": {"text": " Программная система классификации жалоб поступающих в службу 005 с использованием векторной модели представления знаний и машинного обучения", "bold": True, "underline": False},
    "Ф. И. О. субъекта персональных данных": {"text": " Шерстнев Павел Александрович", "bold": True, "underline": True},
    "Адрес места жительства": {"text": " РФ, 662155, г. Ачинск, ул. Кирова, д. 48, кв. 27", "bold": True, "underline": True}
    ""

}

passport_dict = {'passport_series': {'text': "1234", 'bold': True, 'underline': True},
    'passport_number': {'text': "567890", 'bold': True, 'underline': True},
    'issued_by': {'text': "ОУФМС России в г. Москве", 'bold': True, 'underline': True},
    'date': {'text': "01.01.2020", 'bold': True, 'underline': True},
    'division_code': {'text': "123-456", 'bold': True, 'underline': True},}


# Пример использования
file_path = '5.1.Согласие_на_обработку_Шаблон.docx'  # Указывайте путь к вашему .docx файлу
doc = Document(file_path)

append_text_to_docx(doc, append_dict)
add_passport_and_sing(doc, passport_dict, append_dict)


new_file_path = file_path.replace("Шаблон", append_dict["Ф. И. О. субъекта персональных данных"]["text"])
doc.save(new_file_path)
print(f"Документ успешно обновлен и сохранен как {new_file_path}")
