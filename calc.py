import input_data as in_data
import pandas as pd

# Чтение данных из Excel-файла
df = pd.read_excel('database.xlsx')
name_comp = df.iloc[:, 0]
mm_comp = df.iloc[:, 1]
table1_col2 = df.iloc[:, 2]
table1_col3 = df.iloc[:, 3]
table1_col4 = df.iloc[:, 4]
table1_col5 = df.iloc[:, 5]
table1_col6 = df.iloc[:, 6]
table1_col7 = df.iloc[:, 7]
# Преобразование первого столбца в список

name_comp = name_comp.tolist()
mm_comp = mm_comp.tolist()
table1_col2 = table1_col2.tolist()
table1_col3 = table1_col3.tolist()
table1_col4 = table1_col4.tolist()
table1_col5 = table1_col5.tolist()
table1_col6 = table1_col6.tolist()

table1_col1_title = [f'Номер потока по схеме', f'Наименование потока', f'1', 'Cостав'] + [f'{item}' for item in name_comp]
table1_col2_title = [f'1', f'{name_comp[0]} на входе в колонну {in_data.spec_equipment1}',f'2', 'кг/ч'] + [f'{item:.4f}' for item in table1_col2]
table1_col3_title = [None, None,f'3', '% масс.'] + [f'{item:.4f}' for item in table1_col3]
table1_col4_title = [f'2', f'Регенерированный раствор  {name_comp[9]} на входе в колонну {in_data.spec_equipment1}',f'4', 'кг/ч'] + [f'{item:.4f}' for item in table1_col4]
table1_col5_title = [None, None,f'5', '% масс.'] + [f'{item:.4f}' for item in table1_col5]
table1_col6_title = [f'3', f'{name_comp[0]} на выходе из верха колонны {in_data.spec_equipment1}',f'6', 'кг/ч'] + [f'{item:.4f}' for item in table1_col6]
table1_col7_title = [None, None,f'7', '% масс.'] + [f'{item:.4f}' for item in table1_col7]


print(len(table1_col1_title))
print(table1_col1_title)
print(len(table1_col2_title))
print(len(table1_col3_title))
print(len(table1_col4_title))
print(len(table1_col5_title))
print(len(table1_col6_title))
print(len(table1_col7_title))