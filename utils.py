from docxtpl import DocxTemplate
from typing import Dict, List
import os

file_agreements_paths: List[str] = [
    "templates\\Согласие_на_обработку_Шаблон.docx",
    "templates\\Согласие_на_указание_Шаблон.docx",
]

file_report_path: str = "templates\\РЕФЕРАТ_Шаблон.docx"
file_code_path: str = "templates\\Идентифицирующие_ПрЭВМ_Шаблон.docx"


# Утилита для сохранения документа
def save_document(
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
        for author, replacements in replacements_autors.items():
            save_document(
                file_agreement_path,
                replacements | replacements_programm,
                save_path,
                author,
            )


# Реферат
def render_report(
    replacements_autors: Dict[str, Dict[str, str]],
    replacements_programm: Dict[str, str],
    save_path: str,
) -> None:
    author_names = ", ".join(replacements_autors.keys())
    replacements = {"author_names": author_names} | replacements_programm
    save_document(
        file_report_path, replacements, save_path, replacements_programm["theme"]
    )


# Идентифицирующие материалы
def render_code(
    replacements_autors: Dict[str, Dict[str, str]],
    replacements_programm: Dict[str, str],
    save_path: str,
) -> None:
    author_names = ",\n ".join(replacements_autors.keys())
    replacements = {"author_names": author_names} | replacements_programm
    save_document(
        file_code_path, replacements, save_path, replacements_programm["theme"]
    )
