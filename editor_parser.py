from docx import Document
from pathlib import Path
from typing import Dict
import re

def parse_profile_from_docx(filepath: Path) -> Dict[str, Dict[str, str]]:
    doc = Document(filepath)
    data = {}
    data_en = {}

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 2:
                key = cells[0].text.strip()
                value_ru = cells[1].text.strip()
                value_en = cells[2].text.strip() if len(cells) >= 3 else ""
                if key:
                    data[key] = value_ru
                    data_en[key] = value_en

    def normalize(text: str) -> str:
        return re.sub(r"[^a-zа-я0-9]", "", text.lower())

    def is_valid_field(value: str) -> bool:
        lowered = value.lower()
        return (
            value.strip()
            and value.strip() != "-"
            and "найти" not in lowered
            and "важно" not in lowered
            and "профиль" not in lowered
        )

    def get(possible_keys, source):
        for key in source:
            nk = normalize(key)
            for target in possible_keys:
                if normalize(target) in nk:
                    val = source[key].strip()
                    if is_valid_field(val):
                        return val
        return ""

    def clean_text(text: str) -> str:
        return text.replace("-\n", "").replace("\n", " ").strip()

    full_name = get(["ФИО", "Фамилия Имя Отчество"], data)
    if not full_name:
        last = get(["Фамилия"], data)
        first = get(["Имя"], data)
        patronymic = get(["Отчество"], data)
        full_name = f"{last} {first} {patronymic}".strip()

    position = get(["должность в редакции", "редколлегии", "позиция", "членство"], data)
    degree_title = get(["ученое звание", "учёное звание", "ученая степень", "степень", "звание"], data)
    org = get(["основное место работы"], data)
    department = get(["подразделение организации"], data)

    specialization = get(["специализация", "ключевые слова", "область интересов"], data)
    website = get(["личной страницы", "url", "сайт"], data)
    email = get(["электронной почты", "email"], data)
    spin = get(["spin"], data)
    scopus = get(["scopus author id", "scopus id", "scopus"], data)
    researcher = get(["researcher id"], data)
    orcid = get(["orcid"], data)

    # Английские поля
    full_name_en = get(["ФИО", "Фамилия Имя Отчество"], data_en)
    position_en = get(["должность в редакции", "редколлегии", "позиция", "членство"], data_en)
    degree_title_en = get(["ученое звание", "учёное звание", "ученая степень", "степень", "звание"], data_en)
    org_en = get(["основное место работы"], data_en)
    department_en = get(["подразделение организации"], data_en)
    keywords_en = get(["ключевые слова", "keywords", "область интересов"], data_en)

    line1_ru = ", ".join(dict.fromkeys(
        filter(None, [full_name, position, degree_title, org, department])
    ))

    line1_en = ", ".join(dict.fromkeys(
        filter(None, [full_name_en, position_en, degree_title_en, org_en, department_en])
    ))

    return {
        "ru": {
            "ФИО, звание и место работы": clean_text(line1_ru),
            "Специализация": clean_text(specialization),
            "Сайт": clean_text(website),
            "Email": clean_text(email),
            "SPIN": clean_text(spin),
            "Scopus ID": clean_text(scopus),
            "Researcher ID": clean_text(researcher),
            "ORCID": clean_text(orcid)
        },
        "en": {
            "Name, position, affiliation": clean_text(line1_en),
            "Keywords": clean_text(keywords_en)
        }
    }
