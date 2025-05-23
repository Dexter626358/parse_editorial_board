# parse_editorial_board
# Editor Profiles Parser

Скрипт для автоматической обработки анкет членов редакционной коллегии, написанных в формате `.docx`, с целью генерации HTML-контента для сайта журнала на русском и английском языках.

---

## ✨ Возможности

- 📝 Чтение анкет из таблиц Word (2 и 3 колонки)
- 🌍 Вывод англоязычных данных из 3-й колонки: имя, должность, место работы, ключевые слова
- 🔄 Автоматическое формирование HTML-фрагмента в структуре:
  ```html
  <div id="editorialTeam">
    <p style="padding-left: 40px;"><strong>ФИО</strong>, Должность, звание, организация</p>
    ...
  </div>
  ```
- 🏠 Поддержка вывода e-mail, SPIN, Scopus, Researcher ID, ORCID
- 📄 Генерация:
  - `editor_profiles.html` — на русском
  - `editor_profiles_en.html` — на английском
  - `editor_profiles_en.txt` — англоязычные строки для справки

---

## 📂 Структура проекта

```
editor_profiles_parser/
├── editor_parser.py              # Модуль для парсинга .docx
├── main.py                       # Скрипт генерации HTML
├── profiles/                     # Папка с .docx-анкетами
├── output/                       # Результаты
│   ├── editor_profiles.html
│   ├── editor_profiles_en.html
│   └── editor_profiles_en.txt
└── requirements.txt              # Зависимости
```

---

## ⚙️ Установка и запуск

1. Создайте виртуальную среду:
```bash
python -m venv .venv
```
2. Активируйте её:
```bash
.venv\Scripts\activate     # Windows
source .venv/bin/activate   # Linux/macOS
```
3. Установите зависимости:
```bash
pip install -r requirements.txt
```
4. Поместите `.docx` файлы в папку `profiles/`
5. Запустите скрипт:
```bash
python main.py
```

---

## 🔍 Пример вывода

```html
<p style="padding-left: 40px;"><strong>Ivanov Ivan Ivanovich</strong>, Editorial Board Member, Doctor of Sciences, Professor, MGU</p>
<p style="padding-left: 80px;"><strong>Keywords:&nbsp;</strong> Fluid mechanics, turbulence</p>
<p style="padding-left: 80px;"><strong>ORCID: </strong><a href="https://orcid.org/0000-0001-2345-6789">0000-0001-2345-6789</a></p>
```

---

## 🎓 Требования

- Python >= 3.8
- `python-docx`

---

## ✍️ Авторы

- 👤 Пользователь: постановка задачи, анкеты
- 🧠 ChatGPT: автоматизация, кодирование, шаблоны

---

## 🚀 Возможные улучшения

- 📷 Вставка фотографий
- 🔝 Группировка по ролям (редколлегия, главред и пр.)
- 💾 Экспорт в `.csv`, `.docx`, `.pdf`
- 🔄 Интеграция с CMS или системой верстки

