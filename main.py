# gui_payment_request.py

import tkinter as tk
from tkinter import messagebox
from docx import Document
from datetime import datetime

def create_payment_request(employee_name, employee_id, hours_worked, hourly_rate, supervisor_position, supervisor_initials):
    # Создаем новый документ
    doc = Document()

    # Заголовок
    doc.add_heading('Заявление на почасовую оплату', 0)

    # Дата
    doc.add_paragraph(f'Дата: {datetime.now().strftime("%d.%m.%Y")}')

    # Информация о сотруднике
    doc.add_paragraph(f'Сотрудник: {employee_name}')
    doc.add_paragraph(f'Табельный номер: {employee_id}')

    # Информация о почасовой оплате
    doc.add_paragraph(f'Количество отработанных часов: {hours_worked}')
    doc.add_paragraph(f'Почасовая ставка: {hourly_rate:.2f} руб.')

    # Общая сумма к оплате
    total_payment = hours_worked * hourly_rate
    doc.add_paragraph(f'Общая сумма к оплате: {total_payment:.2f} руб.')

    # Подпись сотрудника
    doc.add_paragraph(f'\nПодпись сотрудника: _____________________')

    # Подпись руководителя
    doc.add_paragraph(f'\n{supervisor_position}')
    doc.add_paragraph(f'{supervisor_initials}')
    doc.add_paragraph(f'Подпись руководителя: _____________________')

    # Сохранение документа
    filename = f'Заявление_{employee_id}_{datetime.now().strftime("%Y%m%d")}.docx'
    doc.save(filename)
    print(f'Заявление сохранено как {filename}')
    messagebox.showinfo('Успех', f'Заявление сохранено как {filename}')

def on_submit():
    try:
        employee_name = entry_employee_name.get()
        employee_id = entry_employee_id.get()
        hours_worked = float(entry_hours_worked.get())
        hourly_rate = float(entry_hourly_rate.get())
        supervisor_position = entry_supervisor_position.get()
        supervisor_initials = entry_supervisor_initials.get()

        create_payment_request(employee_name, employee_id, hours_worked, hourly_rate, supervisor_position, supervisor_initials)
    except ValueError:
        messagebox.showerror("Ошибка", "Пожалуйста, введите корректные значения для отработанных часов и почасовой ставки.")

# Создаем окно
root = tk.Tk()
root.title("Заявление на почасовую оплату")

# Поля ввода
tk.Label(root, text="Имя сотрудника").grid(row=0)
entry_employee_name = tk.Entry(root)
entry_employee_name.grid(row=0, column=1)

tk.Label(root, text="Табельный номер").grid(row=1)
entry_employee_id = tk.Entry(root)
entry_employee_id.grid(row=1, column=1)

tk.Label(root, text="Отработанные часы").grid(row=2)
entry_hours_worked = tk.Entry(root)
entry_hours_worked.grid(row=2, column=1)

tk.Label(root, text="Почасовая ставка").grid(row=3)
entry_hourly_rate = tk.Entry(root)
entry_hourly_rate.grid(row=3, column=1)

tk.Label(root, text="Должность руководителя").grid(row=4)
entry_supervisor_position = tk.Entry(root)
entry_supervisor_position.grid(row=4, column=1)

tk.Label(root, text="Руководитель").grid(row=5)
entry_supervisor_initials = tk.Entry(root)
entry_supervisor_initials.grid(row=5, column=1)

# Кнопка отправки
submit_button = tk.Button(root, text="Создать заявление", command=on_submit)
submit_button.grid(row=6, columnspan=2)

# Запуск основного цикла
root.mainloop()
