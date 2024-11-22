import openpyxl
import csv
from datetime import datetime
import re

# Укажите путь к файлу Excel
excel_file = '/Users/andreymitnikov/ALL/Работа/ЕУПСБ/Питон/Задание/herm_data.xlsx'
output_file = '/Users/andreymitnikov/ALL/Работа/ЕУПСБ/Питон/Задание/processed_data_final.csv'

# Функция обработки дат
def process_dates(date):
    """
    Обрабатывает значение из столбца 'Дата'.
    Возвращает три значения: date_from, date_to, notes.
    Если notes заполнено, date_from и date_to остаются пустыми.
    """
    if date is None or str(date).strip() == "":
        return None, None, "Дата отсутствует"

    date = str(date).strip()

    # Проверка точной даты, например, "13-06-1762"
    if '-' in date and date.count('-') == 2:
        try:
            exact_date = datetime.strptime(date, "%d-%m-%Y").date()
            return None, None, exact_date.strftime("%d.%m.%Y")  # Только notes
        except ValueError:
            pass

    # Диапазон дат, например "1785-1790"
    if '-' in date:
        parts = date.split('-')
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            return int(parts[0]), int(parts[1]), None

    # Обработка века и диапазонов, например "нач. XIX в."
    if 'нач' in date and 'XIX' in date:
        return 1800, 1830, None
    if 'кон' in date and 'XIX' in date:
        return 1870, 1900, None
    if 'втор пол' in date and 'XVIII' in date:
        return 1750, 1800, None
    if 'XVIII' in date:
        return 1700, 1800, None
    if 'XVII' in date:
        return 1600, 1700, None

    # Одиночный год
    if date.isdigit():
        year = int(date)
        if 1000 <= year <= 9999:
            return year, year, None

    return None, None, "Дата отсутствует"

# Функция извлечения английского названия
def extract_eng_name(text):
    """
    Извлекает только корректное английское название из строки.
    """
    if text is None:
        return None

    matches = re.findall(r'"([^"]+)"', text)
    eng_names = []
    for match in matches:
        if re.search(r'[a-zA-Z]', match):  # Проверяем, что в строке есть английские буквы
            eng_names.append(match.strip())
    return ', '.join(eng_names) if eng_names else None

# Функция извлечения русского названия
def extract_rus_name(text, eng_name):
    """
    Извлекает русское название, удаляя характеристики и проверяя наличие английского названия.
    """
    if text is None or eng_name:
        return None  # Если есть английское название, русское не заполняем

    # Удаляем характеристики вроде "Карикатура:", "Бытовой тип:"
    text = re.sub(r'^(Карикатура|Бытовой тип|Бытовая сцена):', '', text).strip()
    text = re.sub(r'"[^"]+"', '', text).strip()  # Удаляем текст в кавычках

    if text in {'.', ',', ''} or text.isdigit():  # Исключаем некорректные значения
        return None

    return text

# Обработка данных
processed_data = []
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

for row in sheet.iter_rows(min_row=2, values_only=True):  # Пропускаем заголовок
    acc_num = row[1]  # Учетные номера
    description = row[3]  # Полное описание
    eng_name = extract_eng_name(description)  # Английское название
    rus_name = extract_rus_name(description, eng_name)  # Русское название
    date_from, date_to, notes = process_dates(row[4])  # Даты

    # Если notes заполнено, обнуляем date_from и date_to
    if notes:
        date_from, date_to = None, None

    material = ""
    technique = ""
    if row[5]:  # Материал и техника
        parts = row[5].split(',')
        material = parts[0].strip() if len(parts) > 0 else ""
        technique = parts[1].strip() if len(parts) > 1 else ""

    size = row[6]  # Размеры

    processed_data.append([
        acc_num,
        eng_name,
        rus_name,
        description,
        date_from,
        date_to,
        notes,
        material,
        technique,
        size
    ])

# Сохранение данных в CSV
with open(output_file, mode='w', encoding='utf-8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow([
        'acc_num', 'eng_name', 'rus_name', 'description', 'date_from', 'date_to', 'notes', 'material', 'technique', 'size'
    ])
    writer.writerows(processed_data)

print(f"Обработка завершена! Результат сохранён в '{output_file}'")
