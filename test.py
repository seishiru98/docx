import pandas as pd
from docx import Document


doc = Document()

def read_excel_data(filename, sheet_name):
    try:
        return pd.read_excel(filename, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"ОШИБКА: Файл {filename} не найден.")
    except ValueError:
        print(f"ОШИБКА: Лист {sheet_name} не найден в файле {filename}.")
    except Exception as e:
        print(f"ОШИБКА: Произошла ошибка при чтении файла {filename}: {e}")


# Чтение данных из Excel-файла
df8_1 = read_excel_data('НОРМЫ РАСХОДА ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ.xlsx', '8.1')
df8_2 = read_excel_data('НОРМЫ РАСХОДА ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ.xlsx', '8.2')

# Добавляем таблицу в документ
table = doc.add_table(rows=1, cols=len(df8_1.columns))

# Добавляем заголовки таблицы
hdr_cells = table.rows[0].cells
for i, column_name in enumerate(df8_1.columns):
    hdr_cells[i].text = column_name

# Заполняем таблицу данными из DataFrame
for index, row in df8_1.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        row_cells[i].text = str(value)

