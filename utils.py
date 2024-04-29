from docxtpl import DocxTemplate
from typing import Dict, List
import os
from docx import Document
from docx.shared import Pt


file_agreements_paths: List[str] = [
    "templates\\Согласие_на_обработку_Шаблон.docx",
    "templates\\Согласие_на_указание_Шаблон.docx",
]

file_report_path: str = "templates\\РЕФЕРАТ_Шаблон.docx"
file_code_path: str = "templates\\Идентифицирующие_ПрЭВМ_Шаблон.docx"
file_contract_path: str = "templates\\ДОГОВОР_с_авторами_Шаблон.docx"


def insert_authors_table(document_path, marker_text, authors_info):
    doc = Document(document_path)
    
    # Найти абзац с маркерным текстом
    for paragraph in doc.paragraphs:
        print(marker_text)
        if marker_text in paragraph.text:
            print("нашел")
            # Создание таблицы непосредственно после маркера
            table = doc.add_table(rows=1, cols=3)
            
            # Заполнение заголовков таблицы
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Данные о соавторе, являющемся сотрудником СФУ'
            hdr_cells[1].text = 'Размер, %'
            hdr_cells[2].text = 'Подпись'

            # Заполнение строк таблицы данными авторов
            for author, details in authors_info.items():
                initials = f"{details['subject_name']['name'][0]}.{details['subject_name']['middle_name'][0]}."
                row_cells = table.add_row().cells
                row_cells[0].text = f"{details['subject_name']['surname']} {initials}"
                row_cells[1].text = ''  # Добавьте размер, если он доступен
                row_cells[2].text = '_______________ /'
            
            # Добавить разрыв строки после таблицы
            doc.add_paragraph()
            break  # Прекратить поиск после первого совпадения
    
    # Сохранить измененный документ
    doc.save(document_path)

def format_address(address):
    """
    Форматирует словарь с информацией об адресе в строку в формате паспорта РФ.

    :param address: Словарь с ключами country, post_index, city, street, home_num, appartment_num
    :return: Строка с форматированным адресом
    """
    formatted_address = (
        f"{address['country']}, {address['post_index']}, г. {address['city']}, "
        f"ул. {address['street']}, д. {address['home_num']}, кв. {address['appartment_num']}"
    )
    return formatted_address


def format_author_signature(authors_info):
    """
    Формирует информацию о подписи авторов для документа.
    
    :param authors_info: Словарь с информацией об авторах.
    :return: Строка, содержащая информацию о подписях авторов, разделенная символом новой строки.
    """
    template = "{surname} {name} {middle_name}\n_______________ / {initials} {surname}\n"
    author_texts = []
    for details in authors_info.values():
        initials = f"{details['subject_name']['name'][0]}.{details['subject_name']['middle_name'][0]}."
        text = template.format(
            surname=details['subject_name']['surname'],
            name=details['subject_name']['name'],
            middle_name=details['subject_name']['middle_name'],
            initials=initials
        )
        author_texts.append(text)
    return "\n".join(author_texts)


def format_name(subject_name, full_name=True):
    """
    Форматирует информацию о личности в строку.

    :param subject_name: Словарь с ключами surname, name, middle_name, содержащий личные данные.
    :param full_name: Булевый параметр, который определяет формат вывода:
                      True - полное имя, False - фамилия и инициалы.
    :return: Строка с форматированным именем.
    """
    if full_name:
        # Форматирование полного имени
        return f"{subject_name['surname']} {subject_name['name']} {subject_name['middle_name']}"
    else:
        # Форматирование фамилии и инициалов
        initials = f"{subject_name['name'][0]}.{subject_name['middle_name'][0]}."
        return f"{subject_name['surname']} {initials}"


def format_authors_info(authors_info):
    """
    Форматирует информацию об авторах по заданному шаблону.
    
    :param authors_info: Словарь с информацией об авторах.
    :return: Строка, содержащая информацию обо всех авторах, разделенная символом новой строки.
    """
    template = (
        "{surname}, {name}, {middle_name}, личность удостоверена паспортом серии {series} № {number}, "
        "выданным {issued_by}, код подразделения {division_code}, дата выдачи {issue_date}, "
        "зарегистрированный по адресу {post_index}, г. {city}, дом {home_num}, квартира {appartment_num}, "
        "являющийся (являющаяся) сотрудником Университета, в дальнейшем именуемый «Соавтор»;"
    )
    author_texts = []
    for details in authors_info.values():
        text = template.format(
            surname=details['subject_name']['surname'],
            name=details['subject_name']['name'],
            middle_name=details['subject_name']['middle_name'],
            series=details['passport_series'],
            number=details['passport_number'],
            issued_by=details['passport_issued_by'],
            division_code=details['passport_division_code'],
            issue_date=details['passport_issue_date'],
            post_index=details['subject_address']['post_index'],
            city=details['subject_address']['city'],
            home_num=details['subject_address']['home_num'],
            appartment_num=details['subject_address']['appartment_num']
        )
        author_texts.append(text)
    return "\n".join(author_texts)


# Утилита для сохранения документа
def render_document(
    template_path: str, replacements: Dict[str, str], save_path: str, identifier: str
) -> None:
    doc = DocxTemplate(template_path)
    doc.render(replacements)

    file_name = os.path.basename(template_path.replace("Шаблон", identifier))
    full_file_path = os.path.join(save_path, file_name)
    doc.save(full_file_path)


# Соглашения
def render_agreements(
    replacements_autors: Dict[str, Dict[str, str]],
    replacements_programm: Dict[str, str],
    save_path: str,
) -> None:
    for file_agreement_path in file_agreements_paths:
        for replacements in replacements_autors.values():
            formatted_address = {
                "formatted_address": format_address(replacements["subject_address"])
            }
            formatted_name_short = {
                "subject_name_short": format_name(
                    replacements["subject_name"], full_name=False
                )
            }
            formatted_name_long = {
                "subject_name_long": format_name(replacements["subject_name"])
            }

            render_document(
                file_agreement_path,
                replacements
                | replacements_programm
                | formatted_address
                | formatted_name_short
                | formatted_name_long,
                save_path,
                formatted_name_short["subject_name_short"],
            )


# Реферат
def render_report(
    replacements_autors: Dict[str, Dict[str, str]],
    replacements_programm: Dict[str, str],
    save_path: str,
) -> None:
    author_names = ", ".join(
        format_name(replacements["subject_name"], full_name=True)
        for replacements in replacements_autors.values()
    )
    replacements = {"author_names_long": author_names} | replacements_programm
    render_document(
        file_report_path, replacements, save_path, replacements_programm["theme"]
    )


# Идентифицирующие материалы
def render_code(
    replacements_autors: Dict[str, Dict[str, str]],
    replacements_programm: Dict[str, str],
    save_path: str,
) -> None:

    author_names = "\n ".join(
        format_name(replacements["subject_name"], full_name=False)
        for replacements in replacements_autors.values()
    )

    

    replacements = {"author_names_short": author_names} | replacements_programm
    render_document(
        file_code_path, replacements, save_path, replacements_programm["theme"]
    )


def render_contract(
    replacements_autors: Dict[str, Dict[str, str]],
    replacements_programm: Dict[str, str],
    save_path: str, 
):
    authors_info = {"authors_info" : format_authors_info(replacements_autors)}
    author_signature = {"authors_signature" : format_author_signature(replacements_autors)}
    
    replacements = authors_info | replacements_programm | author_signature

    doc = DocxTemplate(file_contract_path)
    doc.render(replacements)

    file_name = os.path.basename(file_contract_path.replace("Шаблон", replacements_programm["theme"]))
    full_file_path = os.path.join(save_path, file_name)
    doc.save(full_file_path)

    marker_text = "тексттекст"
    
    insert_authors_table(full_file_path, marker_text, replacements_autors)

    doc.save(full_file_path)