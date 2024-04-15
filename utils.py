from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def append_text_to_docx(doc, append_dict):
    
    # Поиск и дописывание текста в документе
    for paragraph in doc.paragraphs:
    
        # Обработка статической информации (ФИО, название программы)
        for key, settings in append_dict.items():
            if key in ['passport_series', 'passport_number', 'issued_by', 'date', 'division_code']:
                continue  # Пропускаем параметры паспорта, т.к. они уже обработаны
            if key in paragraph.text:
                for run in paragraph.runs:
                    if key in run.text:
                        new_run = paragraph.add_run(settings['text'])
                        new_run.font.bold = settings['bold']
                        new_run.font.underline = settings['underline']
                        new_run.font.name = 'Times New Roman'
                        new_run.font.size = run.font.size


def add_passport_and_sing(doc, passport_dict, append_dict):
    #5.1.Согласие_на_обработку_Шаблон.docx


    for paragraph in doc.paragraphs:
        if 'Документ, удостоверяющий личность субъекта персональных данных, дата его выдачи и выдавший орган' in paragraph.text:
            # Проверяем каждый run в параграфе
            for run in paragraph.runs:
                if 'Документ, удостоверяющий личность субъекта персональных данных, дата его выдачи и выдавший орган' in run.text:
                    # Собираем итоговую строку из элементов словаря для паспорта
                    formatted_text = (
                        f" паспорт {passport_dict['passport_series']['text']} {passport_dict['passport_number']['text']}"
                        f" выдан {passport_dict['issued_by']['text']}, {passport_dict['date']['text']}, {passport_dict['division_code']['text']}"
                    )
                    new_run = paragraph.add_run(formatted_text)
                    new_run.font.bold = passport_dict['passport_series']['bold']  # Применяем форматирование полужирным если указано
                    new_run.font.underline = passport_dict['passport_series']['underline']  # Применяем подчеркивание если указано
                    new_run.font.name = 'Times New Roman'
                    new_run.font.size = run.font.size

    # Поиск и дописывание текста в документе
    for paragraph in doc.paragraphs:
        if "Подпись" in paragraph.text:
            # Удаление текущего содержимого параграфа
            for run in paragraph.runs:
                run.clear()

            # Добавление "Подпись" и установка табуляции
            run = paragraph.add_run('Подпись')
            run.font.bold = False
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            tab_stops = paragraph.paragraph_format.tab_stops
            tab_stop = tab_stops.add_tab_stop(Pt(450), alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT)

            # Добавление табуляции и "Фамилия и инициалы"
            run.add_tab()
            run = paragraph.add_run(append_dict["Ф. И. О. субъекта персональных данных"]["text"])
            run.font.bold = append_dict["Ф. И. О. субъекта персональных данных"]["bold"]
            run.font.underline = append_dict["Ф. И. О. субъекта персональных данных"]["underline"]
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)



# from docx import Document
# from docx.shared import Pt

# def replace_text_in_paragraph(paragraph, search_text, replace_text):
#     """Заменяет текст в параграфе, сохраняя стиль первого run, в котором найден текст."""
#     if search_text in paragraph.text:
#         # Создаём временный параграф в памяти для сборки нового текста
#         inline = paragraph.runs
#         # Сохраняем стиль первого run, содержащего искомый текст
#         styles = []
#         for i in range(len(inline)):
#             if search_text in inline[i].text:
#                 styles.append((inline[i].bold, inline[i].underline, inline[i].font.size))
#                 break
        
#         # Заменяем весь текст параграфа с сохранением стилей
#         new_text = paragraph.text.replace(search_text, replace_text)
#         paragraph.clear()
#         run = paragraph.add_run(new_text)
#         if styles:
#             run.bold, run.underline, run.font.size = styles[0]
#         else:
#             run.bold = run.underline = None
#             run.font.size = Pt(12)  # Задаём стандартный размер, если стили не найдены

# def append_text_to_docx(doc, append_dict):
#     for paragraph in doc.paragraphs:
#         for key, settings in append_dict.items():
#             if key not in ['passport_series', 'passport_number', 'issued_by', 'date', 'division_code']:
#                 replace_text_in_paragraph(paragraph, key, settings['text'])

# # Пример использования
# doc = Document('path_to_your_document.docx')
# append_dict = {
#     "Название программы для ЭВМ или базы данных": {"text": "Новое название программы", "bold": True, "underline": False},
#     "Ф. И. О. субъекта персональных данных": {"text": "Иванов И. И.", "bold": False, "underline": True}
# }
# append_text_to_docx(doc, append_dict)
# doc.save('path_to_updated_document.docx')