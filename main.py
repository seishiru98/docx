from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENT
import pandas as pd

import input_data as in_data
import calc

# Создаем новый документ
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

class Counter:
    def __init__(self, start_value, step):
        self.value = start_value
        self.step = step

    def increment(self):
        current_value = self.value
        self.value += self.step
        return current_value

class HeadingCounter(Counter):
    def __init__(self, start_value, paragraph_counter, table_counter, fig_counter):
        super().__init__(start_value, 1)
        self.paragraph_counter = paragraph_counter
        self.table_counter = table_counter
        self.fig_counter = fig_counter

    def increment(self):
        current_value = super().increment()
        self.paragraph_counter.reset(current_value + 0.1)
        self.table_counter.reset(current_value + 0.1)
        self.fig_counter.reset(current_value + 0.1)
        return current_value

class ParagraphCounter(Counter):
    def reset(self, new_start_value):
        self.value = new_start_value

class TableCounter(Counter):
    def reset(self, new_start_value):
        self.value = new_start_value

class FigCounter(Counter):
    def reset(self, new_start_value):
        self.value = new_start_value

# Инициализация счетчиков
n_heading = 1
n_paragraph_start = n_heading + 0.1
n_table_start = n_heading + 0.1
n_fig_start = n_heading + 0.1

par_counter = ParagraphCounter(n_paragraph_start, 0.1)
table_counter = TableCounter(n_table_start, 0.1)
fig_counter = FigCounter(n_fig_start, 0.1)
head_counter = HeadingCounter(n_heading, par_counter, table_counter, fig_counter)

def read_excel_data(filename, sheet_name):
    try:
        return pd.read_excel(filename, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"ОШИБКА: Файл {filename} не найден.")
    except ValueError:
        print(f"ОШИБКА: Лист {sheet_name} не найден в файле {filename}.")
    except Exception as e:
        print(f"ОШИБКА: Произошла ошибка при чтении файла {filename}: {e}")

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_1 = head_counter.increment()

heading = doc.add_heading(f'{ch_1:.0f} ВВЕДЕНИЕ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_2 = head_counter.increment()

heading = doc.add_heading(f'{ch_2:.0f} ОБЩИЕ СВЕДЕНИЯ О ТЕХНОЛОГИИ «ДЕМЕРУС»', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width


ch_3 = head_counter.increment()

heading = doc.add_heading(f'{ch_3:.0f} ТЕХНИЧЕСКИЙ УРОВЕНЬ, ПАТЕНТОСПОСОБНОСТЬ И ПАТЕНТНАЯ ЧИСТОТА ПРОЦЕССА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width


ch_4 = head_counter.increment()

heading = doc.add_heading(f'{ch_4:.0f} ХАРАКТЕРИСТИКА ИСХОДНОГО СЫРЬЯ, ПРОДУКТОВ, ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width


ch_5 = head_counter.increment()

heading = doc.add_heading(f'{ch_5:.0f} ТЕХНИЧЕСКАЯ ХАРАКТЕРИСТИКА ОТХОДОВ И ОТРАБОТАННОГО ВОЗДУХА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
new_section = doc.add_section(WD_SECTION.NEW_PAGE)
new_section.orientation = WD_ORIENT.PORTRAIT

ch_6 = head_counter.increment()

heading = doc.add_heading(f'{ch_6:.0f} ТЕХНОЛОГИЯ ПРОЦЕССА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_7 = head_counter.increment()

heading = doc.add_heading(f'{ch_7:.0f} УСЛОВИЯ ПРОВЕДЕНИЯ ПРОЦЕССА «ДЕМЕРУС»', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
new_section = doc.add_section(WD_SECTION.NEW_PAGE)
new_section.orientation = WD_ORIENT.PORTRAIT

ch_8 = head_counter.increment()

heading = doc.add_heading(f'{ch_8:.0f} НОРМЫ РАСХОДА ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

table8_1 = table_counter.increment()

text = [f'',
        f'',
        f'Таблица {table8_1} – Нормы расхода химреагентов и катализаторов при демеркаптанизации СУГ']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

df8_1 = read_excel_data('НОРМЫ РАСХОДА ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ.xlsx', '8.1')

# Добавляем таблицу в документ
table = doc.add_table(rows=1, cols=len(df8_1.columns))
table.style = 'Table Grid'

# Добавляем заголовки таблицы
hdr_cells = table.rows[0].cells
for i, column_name in enumerate(df8_1.columns):
    cell_paragraph = hdr_cells[i].paragraphs[0]
    cell_paragraph.text = column_name
    cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in cell_paragraph.runs:
        set_font(run, 'Times New Roman', 12)
    set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                         line_spacing=18, space_after=0, space_before=0)

# Заполняем таблицу данными из DataFrame
for index, row in df8_1.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        cell_paragraph = row_cells[i].paragraphs[0]
        cell_paragraph.text = str(value)
        cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in cell_paragraph.runs:
            set_font(run, 'Times New Roman', 12)
        set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=18, space_after=0, space_before=0)

table8_2 = table_counter.increment()

text = [f'',
        f'Таблица {table8_2} – Эксплуатационные расходы энергоресурсов при демеркаптанизации СУГ']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

df8_2 = read_excel_data('НОРМЫ РАСХОДА ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ.xlsx', '8.2')

# Добавляем таблицу в документ
table = doc.add_table(rows=1, cols=len(df8_2.columns))
table.style = 'Table Grid'

# Добавляем заголовки таблицы
hdr_cells = table.rows[0].cells
for i, column_name in enumerate(df8_2.columns):
    cell_paragraph = hdr_cells[i].paragraphs[0]
    cell_paragraph.text = column_name
    cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in cell_paragraph.runs:
        set_font(run, 'Times New Roman', 12)
    set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                         line_spacing=18, space_after=0, space_before=0)

# Заполняем таблицу данными из DataFrame
for index, row in df8_2.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        cell_paragraph = row_cells[i].paragraphs[0]
        cell_paragraph.text = str(value)
        cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in cell_paragraph.runs:
            set_font(run, 'Times New Roman', 12)
        set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=18, space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_9 = head_counter.increment()

heading = doc.add_heading(f'{ch_9} МАТЕРИАЛЬНЫЙ И ТЕПЛОВОЙ БАЛАНС ТЕХНОЛОГИИ «ДЕМЕРУС»', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

ch_9_par_1 = par_counter.increment()

content1 = [
    "",
    "",
    f"{ch_9_par_1:.2} Исходные данные для расчета материального баланса",
    "",
    f"1. Число часов работы в году: {None}",
    f"2. Содержание сернистых соединений в СУГ до очистки:",
    f"- сероводород {None}, % масс.",
    f"- метилмеркаптан по сере {None}, % масс.",
    f"- этилмеркаптан по сере {None}, % масс.",
    f"- пропилмеркаптан по сере {None}, % масс.",
    f"- карбонилсульфид по сере {None}, % масс.",
    f"- сероуглерод {None}, % масс.",
    f"- общая меркаптановая сера {None}, % масс.",
    f"3. Содержание сернистых соединений в СУГ после очистки:",
    f"- сероводород {None}, % масс.",
    f"- метилмеркаптан {None}, % масс.",
    f"- этилмеркаптан {None}, % масс.",
    f"- пропилмеркаптан по сере {None}, % масс.",
    f"- карбонилсульфид по сере {None}, % масс.",
    f"- сероуглерод по сере {None}, % масс.",
    f"- общее содержание меркаптановой серы, не более {None}",
    f"- содержание диметилдисульфида по сере {None}",
    f"- содержание диэтилдисульфида по сере {None}",
    f"- содержание общей серы (не более) {None}",
    f"4. Диапазон устойчивой работы установки {None}%÷{None}%",
    f"Расчет материального баланса произведен из условий работы установки по сырью, равного {None} или {None} кг/ч."
]

ch_9_par_2 = par_counter.increment()

content2 = [
    "",
    f"{ch_9_par_2:.2} Расчетный материальный баланс технологии «ДЕМЕРУС»",
    "",
    f"Результаты расчета материального баланса технологии «ДЕМЕРУС» представлены в табл. {None}, схема материальных потоков блока очистки СУГ по технологии «ДЕМЕРУС» - на рис. {None}.",
    "",
    f"Расчет частоты замены 5% раствора щелочи в емкости {None}",
    "",
    f"Взаимодействие щелочи с углекислым газом протекает по следующей реакции:",
    f"NaOH + CO2 → NaHCO3;",
    f"Расход воздуха – {None} нм3/ч, содержание {None} – {None} % ({None} кг/ч).",
    f"Количество NaOH для полного превращения {None}:",
    f"{None} (кг/ч) ⸱ {None} (г/моль) / {None} (г/моль) = {None} (кг/ч),",
    f"где {None} – молярная масса {None}, {None} – молярная масса {None}.",
    f"Принимаем, что в реакцию вступит 4% NaOH из 5%:",
    f"{None} (кг/ч) ⸱ 5% / 4% = {None} (кг/ч)",
    f"Следовательно, необходимый расход NaOH – {None} кг/ч. Расход 5% раствора NaOH равен:",
    f"{None} / {None} = {None} кг/ч ⸱ {None} = {None} кг/год ",
    f"Объем емкости {None} – {None}, степень заполнения – {None}. Следовательно, емкость {None} вмещает:",
    f"{None} ⸱ {None} = {None} м3 ({None} кг раствора NaOH).",
    f"Частота замены раствора для обеспечения необходимого расхода NaOH:",
    f"{None} / {None} ⁓ {None} раз в год."
]

ch_9_par_3 = par_counter.increment()

content3 = [
    "",
    f"{ch_9_par_3:.2} Тепловой баланс стадии экстракции меркаптанов в {None}",
    "",
    f"Тепловой эффект взаимодействия этилмеркаптана со щелочью составляет {None} кДж/моль [14]. Для расчетов примем тепловой эффект взаимодействия метилмеркаптана со щелочью равным тепловому эффекту взаимодействия этилмеркаптана со щелочью.",
    f"Общее число молей меркаптанов, вступивших в реакцию со щелочью, составляет {None} моль/час:",
    f"{None} кг/ч ⸱ ({None}/{None}) (кг/моль) = {None} (моль/час)",
    f"1. Тепловой эффект реакции взаимодействия меркаптанов со щелочью равен:",
    f"Q1 = {None} ⸱ {None} = {None} кДж/час",
    f"Теплота реакции взаимодействия сероводорода со щелочью составляет – {None} кДж/моль. Общее число молей сероводорода, вступивших в реакцию со щелочью, составляет: {None} моль/час ({None} кг/ч⸱({None}/{None})(кг/моль) = {None}).",
    f"Q2 = {None} ⸱ {None} = {None} кДж/час",
    f"3. Количество теплоты, приходящее с СУГ:",
    f"Q3 = {None} ⸱ {None} = {None} кДж/час",
    f"где: {None} – уд. теплоемкость СУГ, кДж/кг °С;",
    f"4. Количество теплоты, приходящее со щелочным раствором:",
    f"Q4 = {None} ⸱ {None} = {None} кДж/час",
    f"где: {None} – уд. теплоемкость 10%-ного раствора щелочи, кДж/кг °С.",
    f"5. Теплопотери в колонне {None} Q5 принимаем по нормируемой плотности теплового потока qL в соответствии СП на тепловую изоляцию:",
    f"qL= {None} Вт/м;",
    f"Q5= {None} Вт/м ⸱ {None} м = {None} Вт = {None} кДж/ч",
    f"6. Количество теплоты, уносимое с очищенным СУГ:",
    f"Q6 = {None} ⸱ tвых ⸱ {None} = {None} ⸱ tвых",
    f"7. Количество теплоты, уносимое с насыщенным меркаптидами щелочным раствором:",
    f"Q7 = {None} ⸱ tвых ⸱ {None} = {None} ⸱ tвых",
    f"Решая уравнение Q1 + Q2 + Q3 + Q4 - Q5 = Q6 + Q7 относительно tвых, получим:",
    f"tвых= ({None} + {None} + {None} + {None} - {None}) / ({None}  + {None}) = {None}°С.",
    f"Следовательно, при смешении раствора щелочи с температурой {None}°С с СУГ, имеющим температуру {None}°С, температура СУГ в колонне {None} c учетом теплопотерь повысится на {None}℃ до {None} °С."
]


text = content1 + content2 + content3

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width


table9_1 = table_counter.increment()

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

    paragraph_after_break = doc.add_paragraph(
        f'Таблица {table9_1} – Расчетный материальный баланс процесса очистки СУГ')
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                         line_spacing=22, space_after=0, space_before=0)
    columns = [table_name_comp]
    for i in range(len(flow_tables)):
        columns.append(flow_tables[i][0])  # Номер потока
        columns.append(flow_tables[i][1])  # Доля потока

    len_status = all(len(col) == len(columns[0]) for col in columns)
    print(f"Все колонки имеют одинаковое количество элементов:", len_status)

    page_size = 6
    for start in range(1, len(columns), page_size):
        end = start + page_size
        current_columns = [table_name_comp] + columns[start:end]

        if len(current_columns) < page_size + 1:
            for _ in range(page_size + 1 - len(current_columns)):
                current_columns.append([''] * len(table_name_comp))

        table = doc.add_table(rows=len(table_name_comp), cols=len(current_columns))
        table.style = 'Table Grid'

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

        paragraph_after_break = doc.add_paragraph(
            f'Продолжение таблицы {table9_1} – Расчетный материальный баланс процесса очистки СУГ')
        paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in paragraph_after_break.runs:
            set_font(run, 'Times New Roman', 14)
        set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=22, space_after=0, space_before=0)


#-----------------------------------------------------------------------------------------------------------------------
table9_2 = table_counter.increment()

def generate_thermodynamics_tables(df, flow_names, flow_nums):
    thermodynamics_id = td.iloc[:, 0]
    table_thermodynamics_id = ['Номер потока по схеме', 'Наименование потока', 'Показатели'] + thermodynamics_id.tolist()

    thermodynamics_tables = []
    for i in range(len(flow_names)):
        liquid_phase = td.iloc[:, 2 + i * 2].tolist()
        gas_phase = td.iloc[:, 3 + i * 2].tolist()

        table_liquid_phase = [flow_nums[i], flow_names[i], 'Жидкая фаза'] + [f'{item:.4f}' for item in liquid_phase]
        table_gas_phase = ['', '', 'Газовая фаза'] + [f'{item:.4f}' for item in gas_phase]

        thermodynamics_tables.append((table_liquid_phase, table_gas_phase))

    return table_thermodynamics_id, thermodynamics_tables


# Чтение данных из Excel-файла
td = read_excel_data('database.xlsx', 'Термодинамика')
fn = read_excel_data('database.xlsx', 'Название потоков')

if td is not None and fn is not None:
    flow_names = fn.iloc[:, 1].tolist()
    flow_nums = fn.iloc[:, 0].tolist()

    table_thermodynamics_id, thermodynamics_tables = generate_thermodynamics_tables(td, flow_names, flow_nums)

    paragraph_after_break = doc.add_paragraph(
        f'Таблица {table9_2} – Термодинамические характеристики потоков процесса очистки СУГ')
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                         line_spacing=22, space_after=0, space_before=0)
    columns = [table_thermodynamics_id]
    for i in range(len(thermodynamics_tables)):
        columns.append(thermodynamics_tables[i][0])  # Номер потока
        columns.append(thermodynamics_tables[i][1])  # Доля потока

    len_status = all(len(col) == len(columns[0]) for col in columns)
    print(f"Все колонки имеют одинаковое количество элементов:", len_status)

    page_size = 6
    for start in range(1, len(columns), page_size):
        end = start + page_size
        current_columns = [table_thermodynamics_id] + columns[start:end]

        if len(current_columns) < page_size + 1:
            for _ in range(page_size + 1 - len(current_columns)):
                current_columns.append([''] * len(table_thermodynamics_id))

        table = doc.add_table(rows=len(table_thermodynamics_id), cols=len(current_columns))
        table.style = 'Table Grid'

        for row_idx in range(len(table_thermodynamics_id)):
            for col_idx, col_data in enumerate(current_columns):
                cell = table.cell(row_idx, col_idx)
                cell_paragraph = cell.paragraphs[0]
                cell_paragraph.text = str(col_data[row_idx])
                cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                for run in cell_paragraph.runs:
                    set_font(run, 'Times New Roman', 12)
                set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                                     line_spacing=18, space_after=0, space_before=0)

        for i in range(1, len(current_columns) - 1, 2):
            table.cell(0, i).merge(table.cell(0, i + 1))
            table.cell(1, i).merge(table.cell(1, i + 1))

        doc.add_page_break()

        paragraph_after_break = doc.add_paragraph(
            f'Продолжение таблицы {table9_2} – Термодинамические характеристики потоков процесса очистки СУГ')
        paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in paragraph_after_break.runs:
            set_font(run, 'Times New Roman', 14)
        set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=22, space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.LANDSCAPE

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width < new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width


fig9_1 = fig_counter.increment()

paragraph_after_break = doc.add_paragraph(
        f'Рисунок {fig9_1} – Материальные потоки блока очистки СУГ от меркаптанов')
paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in paragraph_after_break.runs:
    set_font(run, 'Times New Roman', 14)
set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                     line_spacing=22, space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.LANDSCAPE

# Убедимся, что размеры страницы корректны для альбомной ориентации
if new_section.page_width < new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_10 = head_counter.increment()

heading = doc.add_heading(f'{ch_10} ФИЗИКО-ХИМИЧЕСКИЕ ОСНОВЫ ПРОЦЕССА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

table10_1 = table_counter.increment()

text = [f'',
        f'',
        f'Таблица {table10_1} – Химизм процесса']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_11 = head_counter.increment()

heading = doc.add_heading(f'{ch_11} СПЕЦИФИКАЦИЯ ОСНОВНОГО ТЕХНОЛОГИЧЕСКОГО ОБОРУДОВАНИЯ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

ch_11_par_1 = par_counter.increment()

text = [f'',
        f'',
        f'{ch_11_par_1} Статическое оборудование']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

table11_1 = par_counter.increment()

text = [f'',
        f'Таблица {table11_1} – Спецификация статического оборудования']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

doc.add_page_break()

# Сохраняем документ
doc.save('Мат баланс.docx')
