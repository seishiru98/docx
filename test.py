import pandas as pd
from docx import Document
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENT

# Создание документа
doc = Document()
section = doc.sections[0]

# Установка полей страницы
section.left_margin = Cm(2)  # Левое поле
section.right_margin = Cm(1)  # Правое поле
section.top_margin = Cm(2)  # Верхнее поле
section.bottom_margin = Cm(2)  # Нижнее поле

# Функция для изменения шрифта и размера шрифта
def set_font(run, font_name, font_size):
    run.font.name = font_name
    run.font.size = Pt(font_size)

    # Это нужно для корректного отображения шрифта на всех платформах
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)


# Функция для установки отступов и междустрочного интервала
def set_paragraph_format(paragraph, left_indent=0, right_indent=0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0):
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = Cm(left_indent)
    paragraph_format.right_indent = Cm(right_indent)
    paragraph_format.first_line_indent = Cm(first_line_indent)
    paragraph_format.line_spacing = Pt(line_spacing)
    paragraph_format.space_after = Cm(space_after)
    paragraph_format.space_before = Cm(space_before)


def read_excel_data(filename, sheet_name):
    try:
        return pd.read_excel(filename, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"ОШИБКА: Файл {filename} не найден.")
    except ValueError:
        print(f"ОШИБКА: Лист {sheet_name} не найден в файле {filename}.")
    except Exception as e:
        print(f"ОШИБКА: Произошла ошибка при чтении файла {filename}: {e}")


def generate_flow_tables(df, flow_names, flow_nums):
    name_comp = df.iloc[:, 0]
    table_name_comp = ['Номер потока по схеме', 'Наименование потока', 'Состав'] + name_comp.tolist()

    flow_tables = []
    for i in range(len(flow_names)):
        flow_rate = df.iloc[:, 2 + i * 2].tolist()
        flow_frac = df.iloc[:, 3 + i * 2].tolist()

        table_flow_rate = [flow_nums[i], flow_names[i], 'кг/ч'] + [f'{item:.4f}' for item in flow_rate]
        table_flow_frac = ['', '', '% масс.'] + [f'{item:.4f}' for item in flow_frac]

        flow_tables.append((table_flow_rate, table_flow_frac))

    return table_name_comp, flow_tables


# Чтение данных из Excel-файла
df = read_excel_data('database.xlsx', 'Сводный мат баланс')
fn = read_excel_data('database.xlsx', 'Название потоков')

if df is not None and fn is not None:
    flow_names = fn.iloc[:, 1].tolist()
    flow_nums = fn.iloc[:, 0].tolist()

    table_name_comp, flow_tables = generate_flow_tables(df, flow_names, flow_nums)

    columns = [table_name_comp]
    for i in range(len(flow_tables)):
        columns.append(flow_tables[i][0])  # Номер потока
        columns.append(flow_tables[i][1])  # Доля потока

    # Проверяем, что все колонки имеют одинаковое количество элементов
    len_status = all(len(col) == len(columns[0]) for col in columns)
    print(f"Все колонки имеют одинаковое количество элементов:", len_status)

    # Разделение на страницы с table_name_comp и по 6 колонок на каждой
    page_size = 6
    for start in range(1, len(columns), page_size):
        end = start + page_size
        current_columns = [table_name_comp] + columns[start:end]

        if len(current_columns) < page_size + 1:
            # Если текущая страница имеет меньше 7 колонок (включая table_name_comp), добавляем пустые колонки
            for _ in range(page_size + 1 - len(current_columns)):
                current_columns.append([''] * len(table_name_comp))

        # Добавляем таблицу с количеством строк, соответствующим размеру table_name_comp
        table = doc.add_table(rows=len(table_name_comp), cols=len(current_columns))
        table.style = 'Table Grid'

        # Заполняем таблицу данными из current_columns
        for row_idx in range(len(table_name_comp)):
            for col_idx, col_data in enumerate(current_columns):
                cell = table.cell(row_idx, col_idx)
                cell_paragraph = cell.paragraphs[0]
                cell_paragraph.text = str(col_data[row_idx])
                cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                for run in cell_paragraph.runs:
                    set_font(run, 'Times New Roman', 12)
                set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                                     line_spacing=18, space_after=0, space_before=0)

        # Объединение ячеек
        for i in range(1, len(current_columns) - 1, 2):
            table.cell(0, i).merge(table.cell(0, i + 1))
            table.cell(1, i).merge(table.cell(1, i + 1))

        doc.add_page_break()

    doc.save('Мат баланс.docx')
else:
    print("ОШИБКА: Не удалось прочитать один или оба листа из файла Excel.")
