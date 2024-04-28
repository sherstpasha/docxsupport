from utils import render_agreements
from utils import render_report
from utils import render_code


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
