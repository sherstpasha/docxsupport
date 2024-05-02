from utils import render_agreements
from utils import render_report
from utils import render_code
from utils import render_contract
from utils import render_form
from utils import render_calculation
from utils import render_notification


replacements_programm = {
    "theme": "Программная система классификации жалоб поступающих в службу 005 с использованием векторной модели представления знаний и машинного обучения",
    "annotation": """Программа для ЭВМ «Программная система формирования рекуррентных нейронных сетей гибридным самоконфигурируемым эволюционным алгоритмом» предназначена для поиска моделей динамических объектов и процессов и позволяет в автоматизированном режиме искусственные конструировать нейронные сети с обратными связями. Для построения модели используется комбинация из самоконфигурируемого алгоритма генетического программирования, который оптимизирует структуру нейронной сети и самоконфигурируемого генетического алгоритма, который оптимизирует весовые коэффициенты нейронной сети. 
Программа для ЭВМ реализована в рамках проекта № 075-15-2022-1121 «Гибридные методы моделирования и оптимизации в сложных системах».
""",
    "computer_type": "Процессор Intel(R) Core(TM) от 1 ГГц, от 128 МБ 03У.",
    "implementation_language": "Python",
    "os_type": "Windows 7/8/10/11",
    "programm_size": "77.2 МБ.",
}

replacements_autors = {
    "author1": {
        "subject_name": {
            "surname": "Шерстнев",
            "name": "Павел",
            "middle_name": "Александрович",
        },
        "subject_address": {
            "country": "РФ",
            "post_index": "662155",
            "city": "Ачинск",
            "street": "Уличная",
            "home_num": "312",
            "appartment_num": "32",
        },
        "passport_series": "1234",
        "passport_number": "567890",
        "passport_issued_by": "ОУФМС России в г. Москве",
        "passport_issue_date": "01.01.2020",
        "passport_division_code": "123-456",
        "birthday": "11",
        "birthmonth": "01",
        "birthyear": "1999",
        "country": "Россия",
        "reason": "в силу свободного",
    },
    "author2": {
        "subject_name": {
            "surname": "Кожин",
            "name": "Константин",
            "middle_name": "Дмитриевич",
        },
        "subject_address": {
            "country": "РФ",
            "post_index": "662155",
            "city": "Красноярск",
            "street": "Ленина",
            "home_num": "22",
            "appartment_num": "33",
        },
        "passport_series": "4321",
        "passport_number": "098765",
        "passport_issued_by": "ОУФМС России в г. Москве",
        "passport_issue_date": "02.02.2010",
        "passport_division_code": "321-654",
        "birthday": "13",
        "birthmonth": "03",
        "birthyear": "2002",
        "country": "Россия",
        "reason": "в силу закона",
    },
}


save_path = ""


render_agreements(
    replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)


render_report(
    replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)


render_code(
    replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)

render_contract(
    replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)


render_form(
    replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)


render_calculation(
    replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)


render_notification(replacements_autors=replacements_autors,
    replacements_programm=replacements_programm,
    save_path=save_path,
)