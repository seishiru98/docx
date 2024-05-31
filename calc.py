import pandas as pd

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







