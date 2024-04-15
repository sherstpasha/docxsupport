from docx import Document
from utils import append_text_to_docx


# Пример использования
file_path = '6.2.Согласие_на_указание_Шаблон.docx'  # Указывайте путь к вашему .docx файлу
doc = Document(file_path)

append_dict = {
    "Название:": {"text": " Программная система классификации жалоб поступающих в службу 005 с использованием векторной модели представления знаний и машинного обучения", "bold": True, "underline": False},
    "Фамилия имя отчество:": {"text": " Шерстнев Павел Александрович", "bold": True, "underline": True},
}


append_text_to_docx(doc, append_dict)


new_file_path = file_path.replace("Шаблон", append_dict["Фамилия имя отчество:"]["text"][1:])
doc.save(new_file_path)
print(f"Документ успешно обновлен и сохранен как {new_file_path}")