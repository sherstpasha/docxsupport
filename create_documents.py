import argparse
import pandas as pd
from utils import render_agreements, render_report, render_code, render_contract, render_form, render_calculation, render_notification

def main(file_path, save_path):
    # Загрузка данных из Excel
    data_excel = pd.read_excel(file_path)
    program_data = data_excel.iloc[0]  # предполагаем, что данные о программе находятся в первой строке

    # Загрузка данных об авторах из второго листа
    authors_data = pd.read_excel(file_path, sheet_name=1)

    # Создание словаря авторов
    replacements_autors = {}
    for index, row in authors_data.iterrows():
        author_key = f"author{index+1}"
        replacements_autors[author_key] = {
            "subject_name": {
                "surname": row['фамилия'],
                "name": row['Имя'],
                "middle_name": row['middle_name'],
            },
            "subject_address": {
                "country": row['Страна'],
                "post_index": row['post_index'],
                "city": row['город'],
                "street": row['улица'],
                "home_num": row['home_num'],
                "appartment_num": row['номер квартиры'],
            },
            "passport_series": row['passport_series'],
            "passport_number": row['passport_number'],
            "passport_issued_by": row['passport_issued_by'],
            "passport_issue_date": row['passport_issue_date'].strftime("%d.%m.%Y"),
            "passport_division_code": row['passport_division_code'],
            "birthday": row['birthday'],
            "birthmonth": row['birthmonth'],
            "birthyear": row['birthyear'],
            "country": row['country'],
            "reason": row['reason'],
        }

    # Обновление словаря для программы
    replacements_programm = {
        "theme": program_data['тема'],
        "annotation": program_data['аннотация'],
        "computer_type": program_data['computer_type'],
        "implementation_language": program_data['implementation_language'],
        "os_type": program_data['os_type'],
        "programm_size": program_data['programm_size'],
    }

    # Генерация документов
    render_agreements(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)
    render_report(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)
    render_code(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)
    render_contract(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)
    render_form(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)
    render_calculation(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)
    # render_notification(replacements_autors=replacements_autors, replacements_programm=replacements_programm, save_path=save_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process Excel data for document generation.")
    parser.add_argument("file_path", type=str, help="Path to the Excel file containing data.")
    parser.add_argument("save_path", type=str, help="Path to save the generated documents.")
    args = parser.parse_args()
    main(args.file_path, args.save_path)
