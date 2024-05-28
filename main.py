from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENT

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

# Инициализация счетчиков
n_heading = 1
n_paragraph_start = n_heading + 0.1
n_table_start = n_heading + 0.1
n_fig_start = n_heading + 0.1

par_counter = ParagraphCounter(n_paragraph_start, 0.1)
table_counter = TableCounter(n_table_start, 0.1)
fig_counter = FigCounter(n_fig_start, 0.1)
head_counter = HeadingCounter(n_heading, par_counter, table_counter, fig_counter)

#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_1 = head_counter.increment()
heading = doc.add_heading(f'{ch_1:.0f} ВВЕДЕНИЕ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_2 = head_counter.increment()
heading = doc.add_heading(f'{ch_2:.0f} ОБЩИЕ СВЕДЕНИЯ О ТЕХНОЛОГИИ «ДЕМЕРУС»', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_3 = head_counter.increment()
heading = doc.add_heading(f'{ch_3:.0f} ТЕХНИЧЕСКИЙ УРОВЕНЬ, ПАТЕНТОСПОСОБНОСТЬ И ПАТЕНТНАЯ ЧИСТОТА ПРОЦЕССА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_4 = head_counter.increment()
heading = doc.add_heading(f'{ch_4:.0f} ХАРАКТЕРИСТИКА ИСХОДНОГО СЫРЬЯ, ПРОДУКТОВ, ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_5 = head_counter.increment()
heading = doc.add_heading(f'{ch_5:.0f} ТЕХНИЧЕСКАЯ ХАРАКТЕРИСТИКА ОТХОДОВ И ОТРАБОТАННОГО ВОЗДУХА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_6 = head_counter.increment()
heading = doc.add_heading(f'{ch_6:.0f} ТЕХНОЛОГИЯ ПРОЦЕССА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_7 = head_counter.increment()
heading = doc.add_heading(f'{ch_7:.0f} УСЛОВИЯ ПРОВЕДЕНИЯ ПРОЦЕССА «ДЕМЕРУС»', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
ch_8 = head_counter.increment()
heading = doc.add_heading(f'{ch_8:.0f} НОРМЫ РАСХОДА ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
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
# Добавление каждого абзаца по отдельности
for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22, space_after=0, space_before=0)

# Добавляем разрыв раздела и изменяем ориентацию на альбомную
new_section = doc.add_section(WD_SECTION.NEW_PAGE)
#-----------------------------------------------------------------------------------------------------------------------
#new_section.orientation = WD_ORIENT.LANDSCAPE
#new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width
#-----------------------------------------------------------------------------------------------------------------------
# Добавляем второй параграф после разрыва раздела
table9_1 = table_counter.increment()
paragraph_after_break = doc.add_paragraph(f'Таблица {table9_1} – Расчетный материальный баланс процесса очистки СУГ')
paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in paragraph_after_break.runs:
    set_font(run, 'Times New Roman', 14)
set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                     line_spacing=22, space_after=0, space_before=0)

# Заполнение таблицы
col_1 = in_data.input_comp
col_2 = calc.col_2
col_3 = calc.col_3
col_4 = calc.col_4
col_5 = calc.col_5
col_6 = calc.col_6
col_7 = calc.col_7

# Проверяем, что все колонки имеют одинаковое количество элементов
columns = [col_1, col_2, col_3, col_4, col_5, col_6, col_7]
len_status = all(len(col) == len(columns[0]) for col in columns)
print(f"Все колонки Таблицы {11} – Расчетный материальный баланс процесса очистки СУГ, имеют одинаковое количество элементов:", len_status)

# Добавляем таблицу с количеством строк, соответствующим размеру col_1
table = doc.add_table(rows=len(col_1), cols=7)
table.style = 'Table Grid'

# Заполняем таблицу данными из col_i
for row_idx in range(len(col_1)):
    for col_idx, col_data in enumerate([col_1, col_2, col_3, col_4, col_5, col_6, col_7]):
        cell = table.cell(row_idx, col_idx)
        cell_paragraph = cell.paragraphs[0]
        cell_paragraph.text = col_data[row_idx]
        cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in cell_paragraph.runs:
            set_font(run, 'Times New Roman', 12)
        set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=18, space_after=0, space_before=0)

# Объединяем ячейки по горизонтали
row_cell_02 = table.cell(0, 1)
row_cell_03 = table.cell(0, 2)
row_cell_02.merge(row_cell_03)

row_cell_12 = table.cell(1, 1)
row_cell_13 = table.cell(1, 2)
row_cell_12.merge(row_cell_13)
#-----------------------------
row_cell_02 = table.cell(0, 3)
row_cell_03 = table.cell(0, 4)
row_cell_02.merge(row_cell_03)

row_cell_14 = table.cell(1, 3)
row_cell_15 = table.cell(1, 4)
row_cell_14.merge(row_cell_15)
#-----------------------------
row_cell_06 = table.cell(0, 5)
row_cell_07 = table.cell(0, 6)
row_cell_06.merge(row_cell_07)

row_cell_16 = table.cell(1, 5)
row_cell_17 = table.cell(1, 6)
row_cell_16.merge(row_cell_17)

# Добавляем разрыв страницы
doc.add_page_break()

paragraph_after_break = doc.add_paragraph(f'Продолжение таблицы {table9_1} – Расчетный материальный баланс процесса очистки СУГ')
paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in paragraph_after_break.runs:
    set_font(run, 'Times New Roman', 14)
set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                     line_spacing=22, space_after=0, space_before=0)

# Заполнение таблицы
col_1 = in_data.input_comp
col_2 = calc.col_2
col_3 = calc.col_3
col_4 = calc.col_4
col_5 = calc.col_5
col_6 = calc.col_6
col_7 = calc.col_7

# Проверяем, что все колонки имеют одинаковое количество элементов
columns = [col_1, col_2, col_3, col_4, col_5, col_6, col_7]
len_status = all(len(col) == len(columns[0]) for col in columns)
print(f"Все колонки Таблицы {11} – Расчетный материальный баланс процесса очистки СУГ, имеют одинаковое количество элементов:", len_status)

# Добавляем таблицу с количеством строк, соответствующим размеру col_1
table = doc.add_table(rows=len(col_1), cols=7)
table.style = 'Table Grid'

# Заполняем таблицу данными из col_i
for row_idx in range(len(col_1)):
    for col_idx, col_data in enumerate([col_1, col_2, col_3, col_4, col_5, col_6, col_7]):
        cell = table.cell(row_idx, col_idx)
        cell_paragraph = cell.paragraphs[0]
        cell_paragraph.text = col_data[row_idx]
        cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in cell_paragraph.runs:
            set_font(run, 'Times New Roman', 12)
        set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=18, space_after=0, space_before=0)

# Объединяем ячейки по горизонтали
row_cell_02 = table.cell(0, 1)
row_cell_03 = table.cell(0, 2)
row_cell_02.merge(row_cell_03)

row_cell_12 = table.cell(1, 1)
row_cell_13 = table.cell(1, 2)
row_cell_12.merge(row_cell_13)
#-----------------------------
row_cell_02 = table.cell(0, 3)
row_cell_03 = table.cell(0, 4)
row_cell_02.merge(row_cell_03)

row_cell_14 = table.cell(1, 3)
row_cell_15 = table.cell(1, 4)
row_cell_14.merge(row_cell_15)
#-----------------------------
row_cell_06 = table.cell(0, 5)
row_cell_07 = table.cell(0, 6)
row_cell_06.merge(row_cell_07)

row_cell_16 = table.cell(1, 5)
row_cell_17 = table.cell(1, 6)
row_cell_16.merge(row_cell_17)

# Добавляем разрыв страницы
doc.add_page_break()

# Сохраняем документ
doc.save('Мат баланс.docx')
