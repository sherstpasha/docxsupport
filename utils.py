from docxtpl import DocxTemplate
from typing import Dict, List
import os


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


file_agreements_paths: List[str] = [
    "templates\\Согласие_на_обработку_Шаблон.docx",
    "templates\\Согласие_на_указание_Шаблон.docx",
]

file_report_path: str = "templates\\РЕФЕРАТ_Шаблон.docx"
file_code_path: str = "templates\\Идентифицирующие_ПрЭВМ_Шаблон.docx"


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
