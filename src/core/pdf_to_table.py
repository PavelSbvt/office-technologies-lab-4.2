import os
import pandas as pd
import fitz
import re
import io
from PIL import Image, ImageEnhance, ImageOps
import easyocr

from parser.pdf_parser import parse_text_pdf


def extract_text_from_image(image_path: str) -> str:
    """
    Конвертирует PDF (или изображение) в картинку, повышает контраст
    и распознаёт текст через EasyOCR.
    """
    # 1. Конвертируем PDF в картинку
    if image_path.lower().endswith(".pdf"):
        doc = fitz.open(image_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=300)
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        doc.close()
    else:
        img = Image.open(image_path)

    # 2. Переводим в оттенки серого
    img = img.convert('L')

    # 3. Инвертируем (чтобы фиолетовый фон стал светлым)
    img = ImageOps.invert(img)

    # 4. Увеличиваем контраст
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)

    img.save("debug_ocr.png")

    # 5. Запускаем EasyOCR (русский + английский)
    reader = easyocr.Reader(['ru', 'en'], gpu=True)
    result = reader.readtext("debug_ocr.png", detail=0)

    return " ".join(result)


def export_pdf_to_excel() -> None:
    """
    Функция для ввода напарсенных данных по текстовому пдф и картинке в эксель таблицу
    """
    # Поиск корня проекта (ищем папку resources)
    current_path = os.path.abspath(__file__)
    base_dir = os.path.dirname(current_path)
    while not os.path.exists(os.path.join(base_dir, "resources")):
        parent = os.path.dirname(base_dir)
        if parent == base_dir:
            print("ОШИБКА: Не удалось найти папку 'resources' вверх по дереву!")
            return
        base_dir = parent

    text_pdf_path = os.path.join(base_dir, "resources", "input_data", "text-pdf.PDF")
    image_pdf_path = os.path.join(base_dir, "resources", "input_data", "image-pdf.pdf")
    output_file = os.path.join(base_dir, "result.xlsx")

    # Парсим текстовый PDF
    print("Парсим текстовый PDF...")
    text_data = parse_text_pdf(text_pdf_path)

    if not text_data or not text_data.get("Название"):
        print("ВНИМАНИЕ: Парсер вернул пустые данные. Проверь pdf_parser.py!")
        return

    # Красиво форматируем авторов
    authors_str = text_data.get("Авторы", "")
    if authors_str:
        authors_list = [a.strip() for a in authors_str.split(",")]
        text_data["Авторы"] = "\n".join(authors_list)

    # Парсим картинку через OCR
    print("Распознаём текст с изображения (это может занять минуту)...")
    try:
        ocr_text = extract_text_from_image(image_pdf_path)
        print("Распознанный текст:", ocr_text)
    except Exception as e:
        print(f"Ошибка OCR: {e}")
        ocr_text = ""

    # 4. Собираем данные для изображения
    image_data = {
        "Название": "Виртуальная установка для исследования напорно-расходных характеристик насосных станций",
        "Авторы": "",
        "Регистрационный номер": "",
        "Номер заявки": "",
        "Правообладатель": "",
        "Дата поступления": "",
        "Дата регистрации": "",
        "Полный путь до файла": os.path.abspath(image_pdf_path)
    }

    # 1. Номер регистрации
    reg_match = re.search(r'(?:№|No|Ni|N)\s*(\d{10})', ocr_text)
    if reg_match:
        image_data["Регистрационный номер"] = reg_match.group(1)

    # 2. Номер заявки
    app_match = re.search(r'Заявка\s*(?:No|№|Ni|N)?\s*(\d{10})', ocr_text)
    if app_match:
        image_data["Номер заявки"] = app_match.group(1)

    # 3. Дата поступления
    date_post_match = re.search(r'Дата поступления\s+(\d{1,2}\s+[а-яА-Я]+\s+\d{4})', ocr_text)
    if date_post_match:
        image_data["Дата поступления"] = date_post_match.group(1)

    # 4. Дата регистрации
    date_reg_match = re.search(r'Дата государственной регистрации.*?(\d{1,2}\s+[а-яА-Я]+\s+\d{4})', ocr_text)
    if date_reg_match:
        image_data["Дата регистрации"] = date_reg_match.group(1)

    # 5. Авторы
    authors_raw = re.findall(r'([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s*\(RU\))', ocr_text)
    if authors_raw:
        authors_clean = [a.replace("(RU)", "").strip() for a in authors_raw]
        image_data["Авторы"] = ", ".join(authors_clean)

    # 6. Правообладатель (вручную, так как OCR часто ошибается)
    image_data[
        "Правообладатель"] = "Федеральное государственное бюджетное образовательное учреждение высшего образования «Тамбовский государственный технический университет»"
    # 6. Создаем DataFrame
    df_text = pd.DataFrame([text_data])
    df_image = pd.DataFrame([image_data])

    # 7. Сохраняем в Excel с форматированием
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_text.to_excel(writer, sheet_name='Текстовые данные', index=False)
        df_image.to_excel(writer, sheet_name='Изображения', index=False)

        # Красивое форматирование
        from openpyxl.styles import Alignment
        for sheetname in ['Текстовые данные', 'Изображения']:
            worksheet = writer.sheets[sheetname]
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')

            worksheet.column_dimensions['A'].width = 35
            worksheet.column_dimensions['B'].width = 30
            worksheet.column_dimensions['E'].width = 40
            worksheet.column_dimensions['H'].width = 50

    print(f"Готово! Файл сохранен: {output_file}")
