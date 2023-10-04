import pandas as pd
from gspread import Cell
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from gspread_formatting import *
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Указать путь к файлу XLS
file_path = 'F:/Work/OutLookConnection/move out 9.15-9.17.xlsx'

# Прочитать файл XLS в DataFrame
moveOut = pd.read_excel(file_path)


def get_objects():
    # Создать словарь для хранения информации об объектах
    objects = {}

    # Пройтись по каждой строке DataFrame и сохранить информацию об объектах в словарь objects
    for _, row in moveOut.iterrows():
        object_no = row['container No']
        # Проверить, содержит ли номер контейнера "RX", если нет, то пропустить этот контейнер
        if not isinstance(object_no, str) or "RX" not in object_no:
            continue
        # Считать значение столбца 'Out Date' как Timestamp
        out_date = row['Out Date']
        # Преобразовать значение Timestamp в строку формата 'дд-мм-гггг'
        formatted_out_date = datetime.strftime(out_date, "%d.%m.%Y")

        object_info = {
            "size": row['size'],
            "type": row['type'],
            "stock": row['stock'],
            "Out Date": formatted_out_date,
            "Owner": row['Owner'],
            "Direction": row['Direction'],
            "Release order": row['Release order'],
            "Storage Days": row['Storage Days']
        }
        objects[object_no] = object_info

    return objects

def get_object_info(container_no, objects):
    # Проверить, есть ли информация о контейнере с заданным номером в словаре objects
    if container_no in objects:
        return objects[container_no]

    return None

# Получить все объекты и сохранить их в переменную
all_objects = get_objects()
print(all_objects)

# Пример использования функции get_object_info()
container_no = 'RXTU4544611'
info = get_object_info(container_no, all_objects)

if info:
    print(f"Информация о контейнере {container_no}:")
    print(info)
else:
    print(f"Контейнер с номером {container_no} не найден.")


# Путь к файлу JSON с ключом для доступа к Google API
json_keyfile = 'F:/Work/OutLookConnection/service_account_key.json'
# ID таблицы Google
spreadsheet_id = '1ubYPsCTwi7cn8r7tUs7zzOUrOoM_2RvDRNDo0Jvu_-I'
# Имя листа в таблице
sheet_name = 'ПО BIC'

# Авторизация с помощью ключа сервисного аккаунта
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
client = gspread.authorize(credentials)

# Открытие таблицы
spreadsheet = client.open_by_key(spreadsheet_id)
# Выбор листа
sheet = spreadsheet.worksheet(sheet_name)


# Функция для поиска номера строки по номеру контейнера
def find_row_number(container_no):
    cell = sheet.find(container_no)
    return cell.row if cell else None
print(find_row_number(container_no))


def container_filler(all_objects):
    cells_to_update = []
    cells_to_format = []

    # Получение всех значений таблицы
    all_values = sheet.get_all_values()

    print(all_values[159][24])

    for container_no, object_info in all_objects.items():
        row_number = find_row_number(container_no)
        if row_number:
            if all_values[row_number - 1][15 - 1] == '':
                cells_to_update.append(Cell(row=row_number, col=15, value=object_info["Out Date"]))
                cells_to_format.append(f'O{row_number}')
                print(f"Ячейка Даты отправления в строке {row_number} были успешно заполнены (2 круг).")
            elif all_values[row_number - 1][25 - 1] == '':
                cells_to_update.append(Cell(row=row_number, col=25, value=object_info["Out Date"]))
                cells_to_format.append(f'Y{row_number}')
                print(f"Ячейка Даты отправления в строке {row_number} были успешно заполнены (3 круг).")
            elif all_values[row_number - 1][35 - 1] == '':
                cells_to_update.append(Cell(row=row_number, col=35, value=object_info["Out Date"]))
                cells_to_format.append(f'AI{row_number}')
                print(f"Ячейка Даты отправления в строке {row_number} были успешно заполнены (4 круг).")
            else:
                print(f"Нет подходящей строки для контейнера {container_no} или ячейка даты уже заполнена.")

            if all_values[row_number - 1][16 - 1] == '':
                cells_to_update.append(Cell(row=row_number, col=16, value=object_info["Direction"]))
                cells_to_format.append(f'P{row_number}')
                print(f"Ячейка Места отправления в строке {row_number} были успешно заполнены (2 круг).")
            elif all_values[row_number - 1][26 - 1] == '':
                cells_to_update.append(Cell(row=row_number, col=26, value=object_info["Direction"]))
                cells_to_format.append(f'Z{row_number}')
                print(f"Ячейка Места отправления в строке {row_number} были успешно заполнены (3 круг).")
            elif all_values[row_number - 1][36 - 1] == '':
                cells_to_update.append(Cell(row=row_number, col=36, value=object_info["Direction"]))
                cells_to_format.append(f'AJ{row_number}')
                print(f"Ячейка Места отправления в строке {row_number} были успешно заполнены (4 круг).")
            print(f"Нет подходящей строки для контейнера {container_no} или ячейка отправления уже заполнена.")

    if len(cells_to_update) > 0:
        sheet.update_cells(cells_to_update)

    for cell in cells_to_format:
        format_cell_range(sheet, cell, CellFormat(backgroundColor=Color(4, 27, 50)))

container_filler(all_objects)
