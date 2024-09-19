import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import psycopg2
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import time



def open_excel_file_T1():
    # Открываем файловый диалог для выбора файла Excel
    file_path_T1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path_T1:
        # Открываем файл Excel с помощью openpyxl
        wb2 = load_workbook(file_path_T1)
        # Выводим сообщение о успешном открытии файла
        label.config(text=f"Файл {file_path_T1} открыт успешно!")

        # Выполняем основной код
        process_excel_file_T1(wb2)

def open_excel_file_T2():
    # Открываем файловый диалог для выбора файла Excel
    file_path_T2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path_T2:
        # Открываем файл Excel с помощью openpyxl
        wb3 = load_workbook(file_path_T2)
        # Выводим сообщение о успешном открытии файла
        label.config(text=f"Файл {file_path_T2} открыт успешно!")

        # Выполняем основной код
        process_excel_file_T2(wb3)




def process_excel_file_T1(wb2):
    # Получаем начальное время
    st = time.time()

    print(wb2.sheetnames)
    sheet_obj = wb2.active
    print("Size:" + str(sheet_obj.max_column) + " X " + str(sheet_obj.max_row))
    count_column = sheet_obj.max_column
    count_row = sheet_obj.max_row
    cell_obj = sheet_obj.cell(row=1, column=1)

    color_arr = ["00000000", "1", "2", "FF8DB4E2", "0", "FFDA9694", "5",0,0,0,0,0]
    print(color_arr[5])

    fooo = [[0] * count_column for i in range(count_row)]
    fooo_color = [[0] * count_column for i in range(count_row)]

    for key, value2 in enumerate(fooo):
        for key1, value1 in enumerate(fooo[key]):
            fooo[key][key1] = str(sheet_obj.cell(row=key + 1, column=key1 + 1).value).strip()
            colorr = str(sheet_obj.cell(row=key + 1, column=key1 + 1).fill.fgColor.rgb)
            if colorr[0] == "V":
                colorr = int(sheet_obj.cell(row=key + 1, column=key1 + 1).fill.fgColor.theme)
                if colorr > 11:
                      colorr = 0
                colorr = color_arr[colorr]
            fooo_color[key][key1] = colorr
    print(fooo[0][3])

    # Подключение к базе данных
    conn = psycopg2.connect(
        dbname="vibory",
        user="elverona",
        password="qwerty",
        host="localhost",
        port="5432"
    )

    # Создание курсора для выполнения операций с базой данных
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS T1")
    command_T1 = ''
    for i in range(count_column):
        command_T1 += 'column' + str(i) + ' VARCHAR(400),'
    command_T1 += 'columnError VARCHAR(100)'
    command_init_T1 = "CREATE TABLE T1 ("+ command_T1 + ")"
    cur.execute(command_init_T1)
    insert_command_T1 = "INSERT INTO T1 VALUES (" + ",".join(["%s"] * count_column) + ", %s)"
    for row in fooo:
        cur.execute(insert_command_T1, row + [None])

    conn.commit()
    conn.close()

    workbook = openpyxl.Workbook()

    # Create a new sheet
    sheet = workbook.active

    # Write the array data to the sheet
    for row_index, row_data in enumerate(fooo_color):
        for column_index, color in enumerate(row_data):
            cell = sheet.cell(row=row_index + 1, column=column_index + 1)
            cell.value = fooo[row_index][column_index]
            if isinstance(color, int):
                color = hex(color)[2:].upper()
            if color != "00000000" and color != "0":
                fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                cell.fill = fill

    # Save the workbook to a file
    workbook.save('C:/1/2.xlsx')

    print("Excel file 'C:/1/2.xlsx' created successfully.")




def process_excel_file_T2(wb3):
    # Получаем начальное время
    st = time.time()

    print(wb3.sheetnames)
    sheet_obj = wb3.active
    print("Size:" + str(sheet_obj.max_column) + " X " + str(sheet_obj.max_row))
    count_column = sheet_obj.max_column
    count_row = sheet_obj.max_row
    cell_obj = sheet_obj.cell(row=1, column=1)

    color_arr = ["00000000", "1", "2", "FF8DB4E2", "0", "FFDA9694", "5", 0, 0, 0, 0, 0]
    print(color_arr[5])

    fooo = [[0] * count_column for i in range(count_row)]
    fooo_color = [[0] * count_column for i in range(count_row)]

    for key, value2 in enumerate(fooo):
        for key1, value1 in enumerate(fooo[key]):
            fooo[key][key1] = str(sheet_obj.cell(row=key + 1, column=key1 + 1).value).strip()
            colorr = str(sheet_obj.cell(row=key + 1, column=key1 + 1).fill.fgColor.rgb)
            if colorr[0] == "V":
                colorr = int(sheet_obj.cell(row=key + 1, column=key1 + 1).fill.fgColor.theme)
                if colorr > 11:
                    colorr = 0
                colorr = color_arr[colorr]
            fooo_color[key][key1] = colorr
    print(fooo[0][3])

    # Подключение к базе данных
    conn = psycopg2.connect(
        dbname="vibory",
        user="elverona",
        password="qwerty",
        host="localhost",
        port="5432"
    )

    # Создание курсора для выполнения операций с базой данных
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS T2")
    command_T2 = ''
    for i in range(count_column):
        command_T2 += 'column' + str(i) + ' VARCHAR(300),'
    command_T2 += 'columnError VARCHAR(100)'
    command_init_T2 = "CREATE TABLE T2 (" + command_T2 + ")"
    cur.execute(command_init_T2)
    insert_command_T2 = "INSERT INTO T2 VALUES (" + ",".join(["%s"] * count_column) + ", %s)"
    for row in fooo:
        cur.execute(insert_command_T2, row + [None])

    conn.commit()
    conn.close()

    workbook = openpyxl.Workbook()

    # Create a new sheet
    sheet = workbook.active

    # Write the array data to the sheet
    for row_index, row_data in enumerate(fooo_color):
        for column_index, color in enumerate(row_data):
            cell = sheet.cell(row=row_index + 1, column=column_index + 1)
            cell.value = fooo[row_index][column_index]
            if isinstance(color, int):
                color = hex(color)[2:].upper()
            if color != "00000000" and color != "0":
                fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                cell.fill = fill

    # Save the workbook to a file
    workbook.save('C:/1/3.xlsx')

    print("Excel file 'C:/1/3.xlsx' created successfully.")

    # получаем время завершения
    et = time.time()

    # считаем время исполнения
    elapsed_time = et - st
    print('Время исполнения:', elapsed_time, 'секунд')




# def create_T3():
#     # Подключение к базе данных
#     conn = psycopg2.connect(
#         dbname="vibory",
#         user="elverona",
#         password="qwerty",
#         host="localhost",
#         port="5432"
#     )
#
#     # Создание курсора для выполнения операций с базой данных
#     cur = conn.cursor()
#
#     # Создание таблицы T3
#     cur.execute("DROP TABLE IF EXISTS T3")
#
#     cur.execute("CREATE TABLE T3 AS SELECT * FROM T1")
#     cur.execute("INSERT INTO T3 SELECT * FROM T2")
# 
#     conn.commit()
#     conn.close()
#
#     label.config(text="Таблица T3 создана успешно!")

def create_T3():
    # Подключение к базе данных
    conn = psycopg2.connect(
        dbname="vibory",
        user="elverona",
        password="qwerty",
        host="localhost",
        port="5432"
    )

    # Создание курсора для выполнения операций с базой данных
    cur = conn.cursor()

    # Создание таблицы T3
    cur.execute("DROP TABLE IF EXISTS T3")

    # Запрос на создание таблицы T3 с данными из T1 и замены из T2
    query = """
        CREATE TABLE T3 AS
        SELECT 
            T1.column1,
            T1.column2,
            T1.column3,
            T1.column4,
            T1.column5,
            T1.column6,
            T1.column7,
            T1.column8,
            T1.column9,
            T1.column10,
            T1.column11,
            T1.column12,
            T2.column1 AS new_column1,
            T2.column6 AS new_column6
        FROM T1
        LEFT JOIN T2 ON T1.column1 = T2.column1 AND T1.column11 = T2.column6
    """
    cur.execute(query)

    # Обновление таблицы T3 с данными из T2
    query = """
        UPDATE T3
        SET column1 = new_column1,
            column6 = new_column6
        WHERE new_column1 IS NOT NULL AND new_column6 IS NOT NULL
    """
    cur.execute(query)

    # Удаление временных столбцов
    query = """
        ALTER TABLE T3
        DROP COLUMN new_column1,
        DROP COLUMN new_column6
    """
    cur.execute(query)

    conn.commit()
    conn.close()

    label.config(text="Таблица T3 создана успешно!")

def create_button(frame, text, command):
    # Создаем кнопку
    button = tk.Button(frame, text=text, command=command)
    button.pack(side=tk.LEFT, padx=5, pady=5)
    return button

# Создаем главное окно
root = tk.Tk()
root.title("Открытие документов Excel")

# Создаем фрейм для кнопок
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Создаем кнопки
button_open_T1 = create_button(frame, "Открыть файл Excel для Т1", open_excel_file_T1)
button_open_T2 = create_button(frame, "Открыть файл Excel для Т2", open_excel_file_T2)
button_create_T3 = create_button(frame, "Создать T3", create_T3)
# button_create_T4 = create_button(frame, "Создать T4", create_T4)
button_exit = create_button(frame, "Выход", root.destroy)

# Создаем метку для вывода сообщений
label = tk.Label(root, text="")
label.pack(padx=10, pady=10)

# Запускаем главное окно
root.mainloop()



