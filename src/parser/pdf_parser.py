import os
import re

import pdfplumber


def parse_text_pdf(pdf_path):
    data = {
        "Название": "",
        "Авторы": "",
        "Регистрационный номер": "",
        "Номер заявки": "",
        "Правообладатель": "",
        "Дата поступления": "",
        "Дата регистрации": "",
        "Полный путь до файла": os.path.abspath(pdf_path)
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text()
    except Exception as e:
        print(f"Ошибка чтения: {e}")
        return data

    flat_text = re.sub(r'\s+', '', text)

    # 1. Регистрационный номер (RU + 10 цифр)
    reg_match = re.search(r'RU(\d{10})', flat_text)
    if reg_match:
        data["Регистрационный номер"] = reg_match.group(1)

    # 2. Дата регистрации
    date_reg_match = re.search(r'Датарегистрации:(\d{2}\.\d{2}\.\d{4})', flat_text)
    if date_reg_match:
        data["Дата регистрации"] = date_reg_match.group(1)

    # 3. Номер заявки и дата поступления (ищем строго после "заявки:", пропуская любые символы до цифр)
    # \D* означает "любые символы, кроме цифр"
    app_match = re.search(r'заявки:\D*(\d{10})\D*(\d{2}\.\d{2}\.\d{4})', flat_text)
    if app_match:
        data["Номер заявки"] = app_match.group(1)
        data["Дата поступления"] = app_match.group(2)

    # 4. Название
    title_match = re.search(r'НазваниепрограммыдляЭВМ:(.*?)Реферат:', flat_text)
    if title_match:
        title = title_match.group(1).replace("Виртуальнаяустановка", "Виртуальная установка ")
        title = title.replace("«Поршневойкомпрессор»", "«Поршневой компрессор»")
        data["Название"] = title.strip()

    # 5. Авторы
    authors_raw = re.findall(r'([А-ЯЁ][а-яё]+[А-ЯЁ][а-яё]+[А-ЯЁ][а-яё]+\(RU\))', flat_text)
    if authors_raw:
        authors_clean = []
        for a in authors_raw:
            # Убираем (RU)
            name = a.replace("(RU)", "")
            # Вставляем пробел перед каждой заглавной буквой (кроме самой первой)
            name_spaced = re.sub(r'(?<!^)(?=[А-ЯЁ])', ' ', name)
            authors_clean.append(name_spaced)
        data["Авторы"] = ", ".join(authors_clean)

    # 6. Правообладатель
    owner_match = re.search(r'Правообладатель\(и\):(.*?)Названиепрограммы', flat_text)
    if owner_match:
        owner = owner_match.group(1).replace("(RU)", "")
        owner = owner.replace("федеральноегосударственное", "федеральное государственное ")
        owner = owner.replace("бюджетноеобразовательное", "бюджетное образовательное ")
        owner = owner.replace("учреждениевысшего", "учреждение высшего ")
        owner = owner.replace("образований«Тамбовский", "образований «Тамбовский ")
        owner = owner.replace("государственныйтехнический", "государственный технический ")
        data["Правообладатель"] = owner.strip()

    return data
