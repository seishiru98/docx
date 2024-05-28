import pandas as pd

# Чтение данных из Excel-файла
df = pd.read_excel('database.xlsx')
first_col = df.iloc[:, 0]
second_col = df.iloc[:, 1]
# Преобразование первого столбца в список

name_comp = first_col.tolist()
mm_comp = second_col.tolist()

# Вывод списка
print(name_comp)
print(mm_comp)
