# СПИСОК КОМПОНЕНТОВ

#Наименование компонентов
#--Имя--Молекулярная масса--Масс концентрация--
name_comp = [
    'СУГ',
    'Нафта',
    'CH3SH',
    'C2H5SH',
    'C3H7SH',
    'H2S',
    'COS',
    'CS2',
    'H2O',
    'NaOH',
    'CH3SNa',
    'C3H7SNa',
    'C2H5SNa',
    'Na2S',
    'CH3SO2SCH3',
    'C2H5SO2SC2H5',
    'C3H7SO2SC3H7',
    'CH3SSCH3',
    'C2H5SSC2H5',
    'C3H7SSC3H7',
    'Na2SO4',
    'N2',
    'O2',
    'CO2'
]
mm_comp =   [
    None,
    None,
    48.108,
    62.136,
    76.162,
    34.081,
    60.075,
    76.141,
    18.015,
    39.997,
    70.090,
    98.144,
    84.117,
    78.045,
    126.198,
    154.252,
    182.306,
    94.200,
    122.254,
    150.308,
    142.041,
    28.014,
    31.998,
    44.009
]

m_conc = []


print(len(name_comp))
print(len(mm_comp))

name_comp1, mm_comp1, m_conc_comp1 = 'СУГ', 55, 55
name_comp2, mm_comp2, m_conc_comp2 = 'Нафта', 55, 55
name_comp3, mm_comp3, m_conc_comp3 = 'CH3SH', 55, 55
name_comp4, mm_comp4, m_conc_comp4 = 'C2H5SH', 55, 55
name_comp5, mm_comp5, m_conc_comp5 = 'C3H7SH', 55, 55
name_comp6, mm_comp6, m_conc_comp6 = 'H2S', 55, 55
name_comp7, mm_comp7, m_conc_comp7 = 'COS', 55, 55
name_comp8, mm_comp8, m_conc_comp8 = 'CS2', 55, 55
name_comp9, mm_comp9, m_conc_comp9 = 'H2O', 55, 55
name_comp10, mm_comp10, m_conc_comp10 = 'NaOH', 55, 55
name_comp11, mm_comp11, m_conc_comp11 = 'CH3SNa', 55, 55
name_comp12, mm_comp12, m_conc_comp12 = 'C2H5SNa', 55, 55
name_comp13, mm_comp13, m_conc_comp13 = 'Na2S', 55, 55
name_comp14, mm_comp14, m_conc_comp14 = 'CH3SO2SCH3', 55, 55
name_comp15, mm_comp15, m_conc_comp15 = 'C2H5SO2SC2H5', 55, 55
name_comp16, mm_comp16, m_conc_comp16 = 'CH3SSCH3', 55, 55
name_comp17, mm_comp17, m_conc_comp17 = 'C2H5SSC2H5', 55, 55
name_comp18, mm_comp18, m_conc_comp18 = 'Na2SO4', 55, 55
name_comp19, mm_comp19, m_conc_comp19 = 'Азот', 55, 55
name_comp20, mm_comp20, m_conc_comp20 = 'Кислород', 55, 55
name_comp21, mm_comp21, m_conc_comp21 = 'СО2', 55, 55

main_flow_rate = 1

# Расход компонентов
flow_comp1 = 16129.5011 # кг/ч
flow_comp2 = 0.0
flow_comp3 = 44.5613
flow_comp4 = 28.3675
flow_comp5 = 0.2269
flow_comp6 = 0.4053
flow_comp7 = 0.9078
flow_comp8 = 0.1621
flow_comp9 = 5.4952
flow_comp10 = 0.0
flow_comp11 = 0.0
flow_comp12 = 0.0
flow_comp13 = 0.0
flow_comp14 = 0.0
flow_comp15 = 0.0
flow_comp16 = 0.0
flow_comp17 = 0.0
flow_comp18 = 0.0
flow_comp19 = 0.0
flow_comp20 = 0.0
flow_comp21 = 0.3728

#----------------------------------------------------------------------------------------------------------------------



# СПЕЦИФИКАЦИЯ ОСНОВНОГО ТЕХНОЛОГИЧЕСКОГО ОБОРУДОВАНИЯ

# Статическое оборудование
spec_equipment1, name_equipment1 = 'C-301', 'Колонна экстракции'
work_temp_equipment1 = [30, 45] # C
work_pres_equipment1 = 20 # кгс/см2
print(f'{spec_equipment1} {name_equipment1}, Рабочая температура {work_temp_equipment1[0]}÷{work_temp_equipment1[1]} С, '
      f'давление {work_pres_equipment1} кгс/см2')

spec_equipment2, name_equipment2 = 'V-301', 'Отстойник СУГ'
work_temp_equipment2 = [30, 45] # C
work_pres_equipment2 = 20 # кгс/см2
print(f'{spec_equipment2} {name_equipment2}, Рабочая температура {work_temp_equipment2[0]}÷{work_temp_equipment2[1]} С, '
      f'давление {work_pres_equipment2} кгс/см2')

spec_equipment3, name_equipment3 = 'R-301', 'Регенератор щелочи'
work_temp_equipment3 = [50, 60] # C
work_pres_equipment3 = [2.35, 4.35] # кгс/см2
print(f'{spec_equipment3} {name_equipment3}, Рабочая температура {work_temp_equipment3[0]}÷{work_temp_equipment3[1]} С, '
      f'давление {work_pres_equipment3[0]}÷{work_pres_equipment3[1]} кгс/см2')

spec_equipment4, name_equipment4 = 'V-302', 'Сепаратор'
work_temp_equipment4 = [50, 60] # C
work_pres_equipment4 = [0.0, 1.0] # кгс/см2
print(f'{spec_equipment4} {name_equipment4}, Рабочая температура {work_temp_equipment4[0]}÷{work_temp_equipment4[1]} С, '
      f'давление {work_pres_equipment4[0]}÷{work_pres_equipment4[1]} кгс/см2')

spec_equipment5, name_equipment5 = 'V-303А', 'Отстойник дисульфидов'
work_temp_equipment5 = [30, 45] # C
work_pres_equipment5 = 12.0 # кгс/см2
print(f'{spec_equipment5} {name_equipment5}, Рабочая температура {work_temp_equipment5[0]}÷{work_temp_equipment5[1]} С,'
      f' давление {work_pres_equipment5} кгс/см2')

spec_equipment6, name_equipment6 = 'V-303B', 'Отстойник дисульфидов'
work_temp_equipment6 = [30, 45] # C
work_pres_equipment6 = 24.0 # кгс/см2
print(f'{spec_equipment6} {name_equipment6}, Рабочая температура {work_temp_equipment6[0]}÷{work_temp_equipment6[1]} С,'
      f' давление {work_pres_equipment6} кгс/см2')

spec_equipment7, name_equipment7 = 'V-304', 'Емкость декарбонизации воздуха'
work_temp_equipment7 = [10, 40] # C
work_pres_equipment7 = [4.0, 6.0] # кгс/см2
print(f'{spec_equipment7} {name_equipment7}, Рабочая температура {work_temp_equipment7[0]}÷{work_temp_equipment7[1]} С,'
      f'давление {work_pres_equipment7[0]}÷{work_pres_equipment7[1]} кгс/см2')

spec_equipment8, name_equipment8 = 'V-305', 'Емкость хранения и приготовления щел. раствора'
work_temp_equipment8 = [10, 70] # C
work_pres_equipment8 = 2.5 # кгс/см2
print(f'{spec_equipment8} {name_equipment8}, Рабочая температура {work_temp_equipment8[0]}÷{work_temp_equipment8[1]} С,'
      f' давление {work_pres_equipment8} кгс/см2')

spec_equipment9, name_equipment9 = 'F-301 А/В', 'Фильтр'
work_temp_equipment9 = [30, 45] # C
work_pres_equipment9 = 21.0 # кгс/см2
print(f'{spec_equipment9} {name_equipment9}, Рабочая температура {work_temp_equipment9[0]}÷{work_temp_equipment9[1]} С,'
      f' давление {work_pres_equipment9} кгс/см2')

spec_equipment10, name_equipment10 = 'F-302 А/В', 'Фильтр'
work_temp_equipment10 = [30, 45] # C
work_pres_equipment10 = 12.0 # кгс/см2
print(f'{spec_equipment10} {name_equipment10}, Рабочая температура {work_temp_equipment10[0]}÷{work_temp_equipment10[1]} С,'
      f'давление {work_pres_equipment10} кгс/см2')

spec_equipment11, name_equipment11 = 'F-303 А/В', 'Фильтр'
work_temp_equipment11 = [30, 45] # C
work_pres_equipment11 = 12.0 # кгс/см2
print(f'{spec_equipment11} {name_equipment11}, Рабочая температура {work_temp_equipment11[0]}÷{work_temp_equipment11[1]} С, '
      f'давление {work_pres_equipment11} кгс/см2')

# Теплообменное оборудование
spec_equipment12, name_equipment12 = 'Е-301', 'Подогреватель насыщенного раствора щелочи'
temp_inlet_equipment12 = 40 # C
temp_outlet_equipment12 = 60 # C
work_pres_equipment12 = 6.0 # кгс/см2
flow_equipment12 = [2.0, 4,0] # м3/ч
print(spec_equipment12, name_equipment12)

spec_equipment13, name_equipment13 = 'ЕW-301', 'Холодильник регенерированного раствора щелочи'
temp_inlet_equipment13 = 60 # C
temp_outlet_equipment13 = 40 # C
work_pres_equipment13 = 12.0 # кгс/см2
flow_equipment13 = [2.0, 4,0] # м3/ч
print(spec_equipment13, name_equipment13)

# Динамическое оборудование
spec_equipment14, name_equipment14 = 'Р-301А/В', 'Насос центробежный'
work_temp_equipment14 = [50, 60] # C
flow_equipment14 = [2.0, 4,0] # м3/ч
print(spec_equipment14, name_equipment14)

spec_equipment15, name_equipment15 = 'Р-302А/В', 'Насос центробежный'
work_temp_equipment15 = [30, 45] # C
flow_equipment15 = [2.0, 4,0] # м3/ч
print(spec_equipment15, name_equipment15)

spec_equipment16, name_equipment16 = 'Р-303', 'Насос центробежный'
work_temp_equipment16 = [10, 70] # C
flow_equipment16 = 10.0 # м3/ч
print(spec_equipment16, name_equipment16)

spec_equipment17, name_equipment17 = 'Р-304А/В', 'Насос плунжерный'
work_temp_equipment17 = 40.0 # C
flow_equipment17 = [0.2, 0.8] # м3/ч
print(spec_equipment17, name_equipment17)
#----------------------------------------------------------------------------------------------------------------------



# Tаблица материального баланса
input_comp = [f'Номер потока по схеме', 'Наименование потока', '1', 'Состав', {name_comp1}, {name_comp2}, {name_comp3},
              {name_comp4}, {name_comp5}, {name_comp6}, {name_comp7}, {name_comp8}, {name_comp9}, {name_comp10},
              {name_comp11}, {name_comp12}, {name_comp13}, {name_comp14}, {name_comp15}, {name_comp16}, {name_comp17},
              {name_comp18}, {name_comp19}, {name_comp20}, {name_comp21}, 'Итого:']

flow_rate1 = 1
print(len(input_comp))





