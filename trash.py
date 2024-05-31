import pandas as pd

# Чтение данных из Excel-файла
df = pd.read_excel('database.xlsx', sheet_name='Сводный мат баланс')
name_comp = df.iloc[:, 0]
mm_comp = df.iloc[:, 1]
execl_C = df.iloc[:, 2]
execl_D = df.iloc[:, 3]
execl_E = df.iloc[:, 4]
execl_F = df.iloc[:, 5]
execl_G = df.iloc[:, 6]
execl_H = df.iloc[:, 7]
execl_I = df.iloc[:, 8]
execl_J = df.iloc[:, 9]
execl_K = df.iloc[:, 10]
execl_L = df.iloc[:, 11]
execl_M = df.iloc[:, 12]
execl_N = df.iloc[:, 13]
execl_O = df.iloc[:, 14]
execl_P = df.iloc[:, 15]
execl_Q = df.iloc[:, 16]
execl_R = df.iloc[:, 17]
execl_S = df.iloc[:, 18]
execl_T = df.iloc[:, 19]
execl_U = df.iloc[:, 20]
execl_V = df.iloc[:, 21]
execl_W = df.iloc[:, 22]
execl_X = df.iloc[:, 23]
execl_Y = df.iloc[:, 24]
execl_Z = df.iloc[:, 25]

fn = pd.read_excel('database.xlsx', sheet_name='Название потоков')
flow_name = fn.iloc[:, 1]
flow_name = flow_name.tolist()
flow_num = fn.iloc[:, 0]
flow_num = flow_num.tolist()

# Преобразование первого столбца в список

name_comp = name_comp.tolist()
mm_comp = mm_comp.tolist()
flow_rate1 = execl_C.tolist()
flow_frac1 = execl_D.tolist()

flow_rate2 = execl_E.tolist()
flow_frac2 = execl_F.tolist()

flow_rate3 = execl_G.tolist()
flow_frac3 = execl_H.tolist()

flow_rate4 = execl_I.tolist()
flow_frac4 = execl_J.tolist()

flow_rate5 = execl_K.tolist()
flow_frac5 = execl_L.tolist()

flow_rate6 = execl_M.tolist()
flow_frac6 = execl_N.tolist()

flow_rate7 = execl_O.tolist()
flow_frac7 = execl_P.tolist()

flow_rate8 = execl_Q.tolist()
flow_frac8 = execl_R.tolist()

flow_rate9 = execl_S.tolist()
flow_frac9 = execl_T.tolist()

flow_rate10 = execl_U.tolist()
flow_frac10 = execl_V.tolist()

flow_rate11 = execl_W.tolist()
flow_frac11 = execl_X.tolist()

flow_rate12 = execl_Y.tolist()
flow_frac12 = execl_Z.tolist()

table_name_comp = [f'Номер потока по схеме', f'Наименование потока', f'1', 'Cостав'] + [f'{item}' for item in name_comp]

table_flow_rate1 = [f'{flow_num[0]}', f'{flow_name[0]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate1]
table_flow_frac1 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac1]

table_flow_rate2 = [f'{flow_num[1]}', f'{flow_name[1]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate2]
table_flow_frac2 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac2]

table_flow_rate3 = [f'{flow_num[2]}', f'{flow_name[2]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate3]
table_flow_frac3 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac3]

table_flow_rate4 = [f'{flow_num[3]}', f'{flow_name[3]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate4]
table_flow_frac4 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac4]

table_flow_rate5 = [f'{flow_num[4]}', f'{flow_name[4]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate5]
table_flow_frac5 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac5]

table_flow_rate6 = [f'{flow_num[5]}', f'{flow_name[5]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate6]
table_flow_frac6 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac6]

table_flow_rate7 = [f'{flow_num[6]}', f'{flow_name[6]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate7]
table_flow_frac7 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac7]

table_flow_rate8 = [f'{flow_num[7]}', f'{flow_name[7]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate8]
table_flow_frac8 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac8]

table_flow_rate9 = [f'{flow_num[8]}', f'{flow_name[8]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate9]
table_flow_frac9 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac9]

table_flow_rate10 = [f'{flow_num[9]}', f'{flow_name[9]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac10 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate11 = [f'{flow_num[10]}', f'{flow_name[10]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac11 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate12 = [f'{flow_num[11]}', f'{flow_name[11]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac12 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]

table_flow_rate13 = [f'{flow_num[12]}', f'{flow_name[12]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac13 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate14 = [f'{flow_num[13]}', f'{flow_name[13]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac14 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate15 = [f'{flow_num[14]}', f'{flow_name[14]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac15 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]

table_flow_rate16 = [f'{flow_num[15]}', f'{flow_name[15]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac16 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate17 = [f'{flow_num[16]}', f'{flow_name[16]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac17 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate18 = [f'{flow_num[17]}', f'{flow_name[17]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac18 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]

table_flow_rate19 = [f'{flow_num[18]}', f'{flow_name[18]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac19 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate20 = [f'{flow_num[19]}', f'{flow_name[19]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac20 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate21 = [f'{flow_num[20]}', f'{flow_name[20]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac21 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]

table_flow_rate22 = [f'{flow_num[21]}', f'{flow_name[21]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac22 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate23 = [f'{flow_num[22]}', f'{flow_name[22]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac23 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate24 = [f'{flow_num[23]}', f'{flow_name[23]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac24 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]

table_flow_rate25 = [f'{flow_num[24]}', f'{flow_name[24]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac25 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate26 = [f'{flow_num[25]}', f'{flow_name[25]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac26 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate27 = [f'{flow_num[26]}', f'{flow_name[26]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac27 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]

table_flow_rate28 = [f'{flow_num[27]}', f'{flow_name[27]}', f'2', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate10]
table_flow_frac28 = [None, None, f'3', '% масс.'] + [f'{item:.4f}' for item in flow_frac10]

table_flow_rate29 = [f'{flow_num[28]}', f'{flow_name[28]}', f'4', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate11]
table_flow_frac29 = [None, None, f'5', '% масс.'] + [f'{item:.4f}' for item in flow_frac11]

table_flow_rate30 = [f'{flow_num[29]}', f'{flow_name[29]}', f'6', 'кг/ч'] + [f'{item:.4f}' for item in flow_rate12]
table_flow_frac30 = [None, None, f'7', '% масс.'] + [f'{item:.4f}' for item in flow_frac12]


paragraph_after_break = doc.add_paragraph(f'Продолжение таблицы {table9_1} – Расчетный материальный баланс процесса очистки СУГ')
paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in paragraph_after_break.runs:
    set_font(run, 'Times New Roman', 14)
set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                     line_spacing=22, space_after=0, space_before=0)
