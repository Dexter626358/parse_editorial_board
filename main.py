from editor_parser import parse_profiles_from_docx
from pathlib import Path
import html
from typing import Optional


def field(label: str, value: str) -> Optional[str]:
    if value and value.strip() and value.strip() != "-":
        return value
    return None


def make_profile_html(profile: dict) -> str:
    html_block = []

    # ФИО жирным, остальное обычным
    full_line = profile["ФИО, звание и место работы"]
    parts = full_line.split(",", 1)
    if len(parts) == 2:
        fio = html.escape(parts[0].strip())
        rest = html.escape(parts[1].strip())
        html_block.append(f'<p style="padding-left: 40px;"><strong>{fio}</strong>, {rest}</p>')
    else:
        html_block.append(f'<p style="padding-left: 40px;"><strong>{html.escape(full_line)}</strong></p>')

    if field("Специализация", profile["Специализация"]):
        html_block.append(f'<p style="padding-left: 80px;"><strong>Специализация:&nbsp;</strong>{html.escape(profile["Специализация"])}</p>')

    if field("Сайт", profile["Сайт"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>Адрес личной страницы в интернете URL: </strong>'
            f'<a href="{html.escape(profile["Сайт"])}" target="_blank" rel="noopener">{html.escape(profile["Сайт"])}</a></p>'
        )

    if field("Email", profile["Email"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>E-mail:</strong> '
            f'<a href="mailto:{html.escape(profile["Email"])} "target="_blank" rel="noopener">{html.escape(profile["Email"])}</a></p>'
        )

    if field("SPIN", profile["SPIN"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>eLibrary SPIN-код: </strong>{html.escape(profile["SPIN"])}</p>'
        )

    if field("Scopus ID", profile["Scopus ID"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>SCOPUS Author ID: </strong>'
           f'<a href="https://www.scopus.com/authid/detail.uri?authorId={html.escape(profile["Scopus ID"])}" target="_blank" rel="noopener">{html.escape(profile["Scopus ID"])}</a></p>'
        )

    if field("Researcher ID", profile["Researcher ID"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>Researcher ID: </strong>'
            f'<a href="https://publons.com/researcher/{html.escape(profile["Researcher ID"])}" target="_blank" rel="noopener">{html.escape(profile["Researcher ID"])}</a></p>'
        )

    if field("ORCID", profile["ORCID"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>ORCID: </strong>'
            f'<a href="https://orcid.org/{html.escape(profile["ORCID"])}" target="_blank" rel="noopener">{html.escape(profile["ORCID"])}</a></p>'
        )

    return "\n".join(html_block)


def main():
    input_dir = Path("profiles")
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    profiles_html = []
    english_lines = []
    english_html = []
    
    # Собираем все профили из всех файлов
    all_profiles = []
    for docx_file in input_dir.glob("*.docx"):
        print(f"Обработка: {docx_file.name}")
        profiles = parse_profiles_from_docx(docx_file)
        print(f"  Найдено анкет: {len(profiles)}")
        all_profiles.extend(profiles)
    
    # Обрабатываем каждый профиль с разделителями
    for idx, profile in enumerate(all_profiles):
        profiles_html.append(make_profile_html(profile["ru"]))
        
        # Добавляем перенос строки между профилями (кроме последнего)
        if idx < len(all_profiles) - 1:
            profiles_html.append('<br>')

        if profile["en"]["Name, position, affiliation"]:
            english_lines.append(profile["en"]["Name, position, affiliation"])

            # HTML блок на английском
            line = profile["en"]["Name, position, affiliation"]
            parts = line.split(",", 1)
            if len(parts) == 2:
                fio = html.escape(parts[0].strip())
                rest = html.escape(parts[1].strip())
                english_html.append(f'<p style="padding-left: 40px;"><strong>{fio}</strong>, {rest}</p>')
            else:
                english_html.append(f'<p style="padding-left: 40px;"><strong>{html.escape(line)}</strong></p>')

            if profile["en"].get("Keywords"):
                english_html.append(
                    f'<p style="padding-left: 80px;"><strong>Keywords:&nbsp;</strong>{html.escape(profile["en"]["Keywords"])}</p>'
                )
        
        # Добавляем перенос строки в английской версии (кроме последнего)
        if idx < len(all_profiles) - 1:
            english_html.append('<br>')

    # Сохраняем HTML на русском
    html_output = ''.join(profiles_html)
    with open(output_dir / "editor_profiles.html", "w", encoding="utf-8") as f:
        f.write(html_output.strip())

    # Сохраняем английский текстовый файл
    with open(output_dir / "editor_profiles_en.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(english_lines))

    # HTML на английском
    english_html_output = ''.join(english_html)
    with open(output_dir / "editor_profiles_en.html", "w", encoding="utf-8") as f:
        f.write(english_html_output.strip())

    print("\n✅ Все HTML-файлы успешно созданы!")


if __name__ == "__main__":
    main()
