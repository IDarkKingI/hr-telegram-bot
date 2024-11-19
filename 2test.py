from docx import Document

def extract_data_from_docx(file_path):
    # Открываем документ
    doc = Document(file_path)
    data = {}
    
    # Перебираем все строки в документе
    for paragraph in doc.paragraphs:
        line = paragraph.text.strip()  # Убираем лишние пробелы
        if ":" in line:
            # Разделяем строку по двоеточию
            key, value = map(str.strip, line.split(":", 1))
            if key and value:  # Проверяем, что ключ и значение не пустые
                data[key] = value  # Сохраняем ключ и значение в словарь
    
    return data

# Путь к файлу
file_path = '/home/darkking/hr-telegram-bot/Карточка самозанятого.docx'

# Извлечение данных
extracted_data = extract_data_from_docx(file_path)

# Печать переменных
if extracted_data:
    for key, value in extracted_data.items():
        print(f"{key}: {value}")
else:
    print("Данные не найдены.")
