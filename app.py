"""
Веб-приложение для обработки zip-архивов с анкетами редакционной коллегии.
"""
import zipfile
import tempfile
from pathlib import Path
from typing import Tuple
from flask import Flask, request, render_template, jsonify
from werkzeug.utils import secure_filename
from editor_parser import parse_profiles_from_docx
import html as html_escape

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB максимум
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()


def field(label: str, value: str) -> bool:
    """Проверяет, есть ли значение в поле."""
    return bool(value and value.strip() and value.strip() != "-")


def make_profile_html(profile: dict) -> str:
    """Создает HTML блок для профиля."""
    html_block = []

    # ФИО жирным, остальное обычным
    full_line = profile["ФИО, звание и место работы"]
    parts = full_line.split(",", 1)
    if len(parts) == 2:
        fio = html_escape.escape(parts[0].strip())
        rest = html_escape.escape(parts[1].strip())
        html_block.append(f'<p style="padding-left: 40px;"><strong>{fio}</strong>, {rest}</p>')
    else:
        html_block.append(f'<p style="padding-left: 40px;"><strong>{html_escape.escape(full_line)}</strong></p>')

    if field("Специализация", profile["Специализация"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>Специализация:&nbsp;</strong>'
            f'{html_escape.escape(profile["Специализация"])}</p>'
        )

    if field("Сайт", profile["Сайт"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>Адрес личной страницы в интернете URL: </strong>'
            f'<a href="{html_escape.escape(profile["Сайт"])}" target="_blank" rel="noopener">'
            f'{html_escape.escape(profile["Сайт"])}</a></p>'
        )

    if field("Email", profile["Email"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>E-mail:</strong> '
            f'<a href="mailto:{html_escape.escape(profile["Email"])}" target="_blank" rel="noopener">'
            f'{html_escape.escape(profile["Email"])}</a></p>'
        )

    if field("SPIN", profile["SPIN"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>eLibrary SPIN-код: </strong>'
            f'{html_escape.escape(profile["SPIN"])}</p>'
        )

    if field("Scopus ID", profile["Scopus ID"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>SCOPUS Author ID: </strong>'
            f'<a href="https://www.scopus.com/authid/detail.uri?authorId={html_escape.escape(profile["Scopus ID"])}" '
            f'target="_blank" rel="noopener">{html_escape.escape(profile["Scopus ID"])}</a></p>'
        )

    if field("Researcher ID", profile["Researcher ID"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>Researcher ID: </strong>'
            f'<a href="https://publons.com/researcher/{html_escape.escape(profile["Researcher ID"])}" '
            f'target="_blank" rel="noopener">{html_escape.escape(profile["Researcher ID"])}</a></p>'
        )

    if field("ORCID", profile["ORCID"]):
        html_block.append(
            f'<p style="padding-left: 80px;"><strong>ORCID: </strong>'
            f'<a href="https://orcid.org/{html_escape.escape(profile["ORCID"])}" '
            f'target="_blank" rel="noopener">{html_escape.escape(profile["ORCID"])}</a></p>'
        )

    return "\n".join(html_block)


def process_zip_archive(zip_path: Path) -> Tuple[str, str]:
    """
    Обрабатывает zip-архив с docx файлами.
    
    Returns:
        Tuple[str, str]: (html_ru, html_en)
    """
    profiles_html = []
    english_html = []
    
    # Создаем временную директорию для распаковки
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Распаковываем архив
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_path)
        
        # Ищем все docx файлы (включая вложенные папки)
        docx_files = list(temp_path.rglob("*.docx"))
        
        if not docx_files:
            raise ValueError("В архиве не найдено файлов .docx")
        
        # Собираем все профили из всех файлов
        all_profiles = []
        for docx_file in docx_files:
            try:
                # Получаем все профили из файла (может быть несколько)
                profiles = parse_profiles_from_docx(docx_file)
                all_profiles.extend(profiles)
            except Exception as e:
                # Пропускаем файлы с ошибками, но логируем
                print(f"Ошибка при обработке {docx_file.name}: {e}")
                continue
        
        # Обрабатываем каждый профиль с разделителями
        for idx, profile in enumerate(all_profiles):
            # Русский HTML
            profiles_html.append(make_profile_html(profile["ru"]))
            
            # Добавляем перенос строки между профилями (кроме последнего)
            if idx < len(all_profiles) - 1:
                profiles_html.append('<br>')
            
            # Английский HTML
            if profile["en"]["Name, position, affiliation"]:
                # HTML блок на английском
                line = profile["en"]["Name, position, affiliation"]
                parts = line.split(",", 1)
                if len(parts) == 2:
                    fio = html_escape.escape(parts[0].strip())
                    rest = html_escape.escape(parts[1].strip())
                    english_html.append(
                        f'<p style="padding-left: 40px;"><strong>{fio}</strong>, {rest}</p>'
                    )
                else:
                    english_html.append(
                        f'<p style="padding-left: 40px;"><strong>{html_escape.escape(line)}</strong></p>'
                    )
                
                if profile["en"].get("Keywords"):
                    english_html.append(
                        f'<p style="padding-left: 80px;"><strong>Keywords:&nbsp;</strong>'
                        f'{html_escape.escape(profile["en"]["Keywords"])}</p>'
                    )
            
            # Добавляем перенос строки между профилями в английской версии (кроме последнего)
            if idx < len(all_profiles) - 1:
                english_html.append('<br>')
    
    html_ru = ''.join(profiles_html).strip()
    html_en = ''.join(english_html).strip()
    
    return html_ru, html_en


@app.route('/')
def index():
    """Главная страница с формой загрузки."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Обработка загруженного zip-архива."""
    if 'file' not in request.files:
        return jsonify({'error': 'Файл не найден'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400
    
    if not file.filename.lower().endswith('.zip'):
        return jsonify({'error': 'Файл должен быть в формате .zip'}), 400
    
    # Сохраняем файл во временную директорию
    filename = secure_filename(file.filename)
    temp_file_path = Path(app.config['UPLOAD_FOLDER']) / filename
    
    try:
        file.save(str(temp_file_path))
        
        # Обрабатываем архив
        html_ru, html_en = process_zip_archive(temp_file_path)
        
        # Удаляем временный файл
        temp_file_path.unlink()
        
        # Возвращаем результаты
        return render_template(
            'result.html',
            html_ru=html_ru,
            html_en=html_en
        )
    
    except zipfile.BadZipFile:
        return jsonify({'error': 'Некорректный zip-архив'}), 400
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'Ошибка обработки: {str(e)}'}), 500
    finally:
        # Убеждаемся, что временный файл удален
        if temp_file_path.exists():
            temp_file_path.unlink()


@app.route('/test_file')
def test_file():
    """Тестовый маршрут для проверки файла с несколькими анкетами."""
    from editor_parser import parse_profiles_from_docx
    
    # Пробуем найти файл
    file_paths = [
        Path("Несколько анкет в одной.docx"),
        Path("..") / "Несколько анкет в одной.docx",
        Path(__file__).parent.parent / "Несколько анкет в одной.docx",
    ]
    
    file_path = None
    for path in file_paths:
        if path.exists():
            file_path = path
            break
    
    if not file_path:
        return jsonify({
            'error': 'Файл "Несколько анкет в одной.docx" не найден',
            'checked_paths': [str(p) for p in file_paths]
        }), 404
    
    try:
        profiles = parse_profiles_from_docx(file_path)
        
        result = {
            'file_path': str(file_path.absolute()),
            'profiles_count': len(profiles),
            'profiles': []
        }
        
        for i, profile in enumerate(profiles, 1):
            profile_info = {
                'number': i,
                'ru': {
                    'fio': profile["ru"].get("ФИО, звание и место работы", ""),
                    'specialization': profile["ru"].get("Специализация", ""),
                    'email': profile["ru"].get("Email", ""),
                    'website': profile["ru"].get("Сайт", ""),
                    'spin': profile["ru"].get("SPIN", ""),
                    'scopus': profile["ru"].get("Scopus ID", ""),
                    'orcid': profile["ru"].get("ORCID", ""),
                },
                'en': {
                    'name_affiliation': profile["en"].get("Name, position, affiliation", ""),
                    'keywords': profile["en"].get("Keywords", ""),
                }
            }
            result['profiles'].append(profile_info)
        
        return jsonify(result)
    
    except Exception as e:
        import traceback
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc()
        }), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

