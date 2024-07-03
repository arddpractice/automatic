from docx import Document

def generate_hourly_payment_document(template_path, name, position, hours_worked, hourly_rate, date):
    # Открытие шаблона документа
    doc = Document(template_path)
    
    # Замена переменных в документе
    for paragraph in doc.paragraphs:
        for key, value in {
            "{{name}}": name,
            "{{position}}": position,
            "{{hours_worked}}": str(hours_worked),
            "{{hourly_rate}}": str(hourly_rate),
            "{{date}}": date
        }.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    
    # Сохранение нового документа
    doc_name = f'Заявление_{name}.docx'
    doc.save(doc_name)
    print(f'Документ сохранен как {doc_name}')

# Пример использования функции
name = input("Введите ваше имя: ")
position = input("Введите вашу должность: ")
hours_worked = int(input("Введите количество отработанных часов: "))
hourly_rate = float(input("Введите почасовую ставку: "))
date = input("Введите дату (дд.мм.гггг): ")

generate_hourly_payment_document('template.docx', name, position, hours_worked, hourly_rate, date)
