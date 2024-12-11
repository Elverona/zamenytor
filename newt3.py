import tkinter as tk
from tkinter import ttk


# Функция для создания чекбоксов
def create_checkboxes(frame, row, col, count):
    var_list = []
    for i in range(count):
        var = tk.BooleanVar()
        cb = tk.Checkbutton(frame, variable=var)
        cb.grid(row=row, column=col + i)
        var_list.append(var)
    return var_list


# Функция для создания выпадающего списка
def create_combobox(frame, row, col, options):
    combo = ttk.Combobox(frame, values=options)
    combo.grid(row=row, column=col)
    return combo


# Создание основного окна
root = tk.Tk()
root.title("Таблица")

# Создание фрейма для таблицы
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Заголовки столбцов
tk.Label(frame, text="Таблица 1").grid(row=0, column=0)
tk.Label(frame, text="__SPK__").grid(row=0, column=1, columnspan=3)
tk.Label(frame, text="Таблица 2").grid(row=0, column=4)

# Данные для таблицы
data = ["Данные 1", "Данные 2", "Данные 3"]

# Заполнение таблицы
for i, item in enumerate(data):
    tk.Label(frame, text=item).grid(row=i + 1, column=0)

    # Чекбоксы по три штуки в столбце __SPK__
    create_checkboxes(frame, i + 1, 1, 3)

    # Выпадающий список в столбце "Таблица 2"
    create_combobox(frame, i + 1, 4, data)

# Строка "Добавить столбец"
tk.Label(frame, text="Добавить столбец").grid(row=len(data) + 1, column=0)

# Один чекбокс в столбце __SPK__ для "Добавить столбец"
var_single = tk.BooleanVar()
cb_single = tk.Checkbutton(frame, variable=var_single)
cb_single.grid(row=len(data) + 1, column=1)

# Выпадающий список в столбце "Таблица 2" для "Добавить столбец"
create_combobox(frame, len(data) + 1, 4, data)

# Запуск основного цикла
root.mainloop()