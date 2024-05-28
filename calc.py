import input_data
import input_data as in_data

# РАСЧЕТ МАТЕРИАЛЬНОГО БАЛАНСА
table1_col2_row4 = in_data.flow_comp1
table1_col2_row5 = in_data.flow_comp2
table1_col2_row6 = in_data.flow_comp3
table1_col2_row7 = in_data.flow_comp4
table1_col2_row8 = in_data.flow_comp5
table1_col2_row9 = in_data.flow_comp6
table1_col2_row10 = in_data.flow_comp7
table1_col2_row11 = in_data.flow_comp8
table1_col2_row12 = in_data.flow_comp9
table1_col2_row13 = in_data.flow_comp10
table1_col2_row14 = in_data.flow_comp11
table1_col2_row15 = in_data.flow_comp12
table1_col2_row16 = in_data.flow_comp13
table1_col2_row17 = in_data.flow_comp14
table1_col2_row18 = in_data.flow_comp15
table1_col2_row19 = in_data.flow_comp16
table1_col2_row20 = in_data.flow_comp17
table1_col2_row21 = in_data.flow_comp18
table1_col2_row22 = in_data.flow_comp19
table1_col2_row23 = in_data.flow_comp20
table1_col2_row24 = in_data.flow_comp21

summ_col2 = (
    table1_col2_row4 + table1_col2_row5 + table1_col2_row6 + table1_col2_row7 + table1_col2_row8 +
    table1_col2_row9 + table1_col2_row10 + table1_col2_row11 + table1_col2_row12 + table1_col2_row13 +
    table1_col2_row14 + table1_col2_row15 + table1_col2_row16 + table1_col2_row17 + table1_col2_row18 +
    table1_col2_row19 + table1_col2_row20 + table1_col2_row21 + table1_col2_row22 + table1_col2_row23 +
    table1_col2_row24
)

table1_col2_row25 = summ_col2

col_2 = [
    f'1',
    f'{in_data.name_comp1} на входе в колонну {in_data.spec_equipment1}',
    f'2',
    'кг/ч',
    f'{table1_col2_row4:.4f}',
    f'{table1_col2_row5:.4f}',
    f'{table1_col2_row6:.4f}',
    f'{table1_col2_row7:.4f}',
    f'{table1_col2_row8:.4f}',
    f'{table1_col2_row9:.4f}',
    f'{table1_col2_row10:.4f}',
    f'{table1_col2_row11:.4f}',
    f'{table1_col2_row12:.4f}',
    f'{table1_col2_row13:.4f}',
    f'{table1_col2_row14:.4f}',
    f'{table1_col2_row15:.4f}',
    f'{table1_col2_row16:.4f}',
    f'{table1_col2_row17:.4f}',
    f'{table1_col2_row18:.4f}',
    f'{table1_col2_row19:.4f}',
    f'{table1_col2_row20:.4f}',
    f'{table1_col2_row21:.4f}',
    f'{table1_col2_row22:.4f}',
    f'{table1_col2_row23:.4f}',
    f'{table1_col2_row24:.4f}',
    f'{table1_col2_row25:.4f}'
]


table1_col3_row4 = table1_col2_row4 / table1_col2_row25 * 100
table1_col3_row5 = table1_col2_row5 / table1_col2_row25 * 100
table1_col3_row6 = table1_col2_row6 / table1_col2_row25 * 100
table1_col3_row7 = table1_col2_row7 / table1_col2_row25 * 100
table1_col3_row8 = table1_col2_row8 / table1_col2_row25 * 100
table1_col3_row9 = table1_col2_row9 / table1_col2_row25 * 100
table1_col3_row10 = table1_col2_row10 / table1_col2_row25 * 100
table1_col3_row11 = table1_col2_row11 / table1_col2_row25 * 100
table1_col3_row12 = table1_col2_row12 / table1_col2_row25 * 100
table1_col3_row13 = table1_col2_row13 / table1_col2_row25 * 100
table1_col3_row14 = table1_col2_row14 / table1_col2_row25 * 100
table1_col3_row15 = table1_col2_row15 / table1_col2_row25 * 100
table1_col3_row16 = table1_col2_row16 / table1_col2_row25 * 100
table1_col3_row17 = table1_col2_row17 / table1_col2_row25 * 100
table1_col3_row18 = table1_col2_row18 / table1_col2_row25 * 100
table1_col3_row19 = table1_col2_row19 / table1_col2_row25 * 100
table1_col3_row20 = table1_col2_row20 / table1_col2_row25 * 100
table1_col3_row21 = table1_col2_row21 / table1_col2_row25 * 100
table1_col3_row22 = table1_col2_row22 / table1_col2_row25 * 100
table1_col3_row23 = table1_col2_row23 / table1_col2_row25 * 100
table1_col3_row24 = table1_col2_row24 / table1_col2_row25 * 100

summ_col3 = (
    table1_col3_row4 + table1_col3_row5 + table1_col3_row6 + table1_col3_row7 + table1_col3_row8 +
    table1_col3_row9 + table1_col3_row10 + table1_col3_row11 + table1_col3_row12 + table1_col3_row13 +
    table1_col3_row14 + table1_col3_row15 + table1_col3_row16 + table1_col3_row17 + table1_col3_row18 +
    table1_col3_row19 + table1_col3_row20 + table1_col3_row21 + table1_col3_row22 + table1_col3_row23 +
    table1_col3_row24
)
table1_col3_row25 = summ_col3


col_3 = [
    None,
    None,
    '3',
    '%',
    f'{table1_col3_row4:.4f}',
    f'{table1_col3_row5:.4f}',
    f'{table1_col3_row6:.4f}',
    f'{table1_col3_row7:.4f}',
    f'{table1_col3_row8:.4f}',
    f'{table1_col3_row9:.4f}',
    f'{table1_col3_row10:.4f}',
    f'{table1_col3_row11:.4f}',
    f'{table1_col3_row12:.4f}',
    f'{table1_col3_row13:.4f}',
    f'{table1_col3_row14:.4f}',
    f'{table1_col3_row15:.4f}',
    f'{table1_col3_row16:.4f}',
    f'{table1_col3_row17:.4f}',
    f'{table1_col3_row18:.4f}',
    f'{table1_col3_row19:.4f}',
    f'{table1_col3_row20:.4f}',
    f'{table1_col3_row21:.4f}',
    f'{table1_col3_row22:.4f}',
    f'{table1_col3_row23:.4f}',
    f'{table1_col3_row24:.4f}',
    f'{table1_col3_row25:.4f}'
]

table1_col4_row4 = 0
table1_col4_row5 = 0
table1_col4_row6 = 0
table1_col4_row7 = 0
table1_col4_row8 = 0
table1_col4_row9 = 0
table1_col4_row10 = 0
table1_col4_row11 = 0
table1_col4_row12 = 0
table1_col4_row13 = 0
table1_col4_row14 = 0
table1_col4_row15 = 0
table1_col4_row16 = 0
table1_col4_row17 = 0
table1_col4_row18 = 0
table1_col4_row19 = 0
table1_col4_row20 = 0
table1_col4_row21 = 0
table1_col4_row22 = 0
table1_col4_row23 = 0
table1_col4_row24 = 0

summ_col4 = (
    table1_col4_row4 + table1_col4_row5 + table1_col4_row6 + table1_col4_row7 + table1_col4_row8 +
    table1_col4_row9 + table1_col4_row10 + table1_col4_row11 + table1_col4_row12 + table1_col4_row13 +
    table1_col4_row14 + table1_col4_row15 + table1_col4_row16 + table1_col4_row17 + table1_col4_row18 +
    table1_col4_row19 + table1_col4_row20 + table1_col4_row21 + table1_col4_row22 + table1_col4_row23 +
    table1_col4_row24
)
table1_col4_row25 = summ_col4

col_4 = [
    f'2',
    f'{in_data.name_comp10} на входе в колонну {in_data.spec_equipment1}',
    f'4',
    'кг/ч',
    f'{table1_col4_row4:.4f}',
    f'{table1_col4_row5:.4f}',
    f'{table1_col4_row6:.4f}',
    f'{table1_col4_row7:.4f}',
    f'{table1_col4_row8:.4f}',
    f'{table1_col4_row9:.4f}',
    f'{table1_col4_row10:.4f}',
    f'{table1_col4_row11:.4f}',
    f'{table1_col4_row12:.4f}',
    f'{table1_col4_row13:.4f}',
    f'{table1_col4_row14:.4f}',
    f'{table1_col4_row15:.4f}',
    f'{table1_col4_row16:.4f}',
    f'{table1_col4_row17:.4f}',
    f'{table1_col4_row18:.4f}',
    f'{table1_col4_row19:.4f}',
    f'{table1_col4_row20:.4f}',
    f'{table1_col4_row21:.4f}',
    f'{table1_col4_row22:.4f}',
    f'{table1_col4_row23:.4f}',
    f'{table1_col4_row24:.4f}',
    f'{table1_col4_row25:.4f}'
]

table1_col5_row4 = table1_col4_row4 / table1_col4_row25 * 100
table1_col5_row5 = table1_col4_row5 / table1_col4_row25 * 100
table1_col5_row6 = table1_col4_row6 / table1_col4_row25 * 100
table1_col5_row7 = table1_col4_row7 / table1_col4_row25 * 100
table1_col5_row8 = table1_col4_row8 / table1_col4_row25 * 100
table1_col5_row9 = table1_col4_row9 / table1_col4_row25 * 100
table1_col5_row10 = table1_col4_row10 / table1_col4_row25 * 100
table1_col5_row11 = table1_col4_row11 / table1_col4_row25 * 100
table1_col5_row12 = table1_col4_row12 / table1_col4_row25 * 100
table1_col5_row13 = table1_col4_row13 / table1_col4_row25 * 100
table1_col5_row14 = table1_col4_row14 / table1_col4_row25 * 100
table1_col5_row15 = table1_col4_row15 / table1_col4_row25 * 100
table1_col5_row16 = table1_col4_row16 / table1_col4_row25 * 100
table1_col5_row17 = table1_col4_row17 / table1_col4_row25 * 100
table1_col5_row18 = table1_col4_row18 / table1_col4_row25 * 100
table1_col5_row19 = table1_col4_row19 / table1_col4_row25 * 100
table1_col5_row20 = table1_col4_row20 / table1_col4_row25 * 100
table1_col5_row21 = table1_col4_row21 / table1_col4_row25 * 100
table1_col5_row22 = table1_col4_row22 / table1_col4_row25 * 100
table1_col5_row23 = table1_col4_row23 / table1_col4_row25 * 100
table1_col5_row24 = table1_col4_row24 / table1_col4_row25 * 100



summ_col5 = (
    table1_col5_row4 + table1_col5_row5 + table1_col5_row6 + table1_col5_row7 + table1_col5_row8 +
    table1_col5_row9 + table1_col5_row10 + table1_col5_row11 + table1_col5_row12 + table1_col5_row13 +
    table1_col5_row14 + table1_col5_row15 + table1_col5_row16 + table1_col5_row17 + table1_col5_row18 +
    table1_col5_row19 + table1_col5_row20 + table1_col5_row21 + table1_col5_row22 + table1_col5_row23 +
    table1_col5_row24
)
table1_col5_row25 = summ_col5

col_5 = [
    None,
    None,
    '5',
    '%',
    f'{table1_col5_row4:.4f}',
    f'{table1_col5_row5:.4f}',
    f'{table1_col5_row6:.4f}',
    f'{table1_col5_row7:.4f}',
    f'{table1_col5_row8:.4f}',
    f'{table1_col5_row9:.4f}',
    f'{table1_col5_row10:.4f}',
    f'{table1_col5_row11:.4f}',
    f'{table1_col5_row12:.4f}',
    f'{table1_col5_row13:.4f}',
    f'{table1_col5_row14:.4f}',
    f'{table1_col5_row15:.4f}',
    f'{table1_col5_row16:.4f}',
    f'{table1_col5_row17:.4f}',
    f'{table1_col5_row18:.4f}',
    f'{table1_col5_row19:.4f}',
    f'{table1_col5_row20:.4f}',
    f'{table1_col5_row21:.4f}',
    f'{table1_col5_row22:.4f}',
    f'{table1_col5_row23:.4f}',
    f'{table1_col5_row24:.4f}',
    f'{table1_col5_row25:.4f}'
]

nt = 1  # Кол-во теоретических тарелок

for _ in range(nt):
    table1_col2_row6 *= 5/100
    table1_col2_row7 *= 10/100
    table1_col2_row8 *= 30/100
    table1_col2_row9 *= 0.0
    table1_col2_row10 *= 70/100
    table1_col2_row11 *= 1.0
    table1_col2_row12 *= 0.0015/100
    table1_col2_row13 *= 0.05/100
    table1_col2_row24 *= 0.0


table1_col6_row4 = table1_col2_row4
table1_col6_row5 = table1_col2_row5
table1_col6_row6 = table1_col2_row6
table1_col6_row7 = table1_col2_row7
table1_col6_row8 = table1_col2_row8
table1_col6_row9 = table1_col2_row9
table1_col6_row10 = table1_col2_row10
table1_col6_row11 = table1_col2_row11
table1_col6_row12 = table1_col2_row12
table1_col6_row13 = table1_col2_row13
table1_col6_row14 = table1_col2_row14
table1_col6_row15 = table1_col2_row15
table1_col6_row16 = table1_col2_row16
table1_col6_row17 = table1_col2_row17
table1_col6_row18 = table1_col2_row18
table1_col6_row19 = table1_col2_row19
table1_col6_row20 = table1_col2_row20
table1_col6_row21 = table1_col2_row21
table1_col6_row22 = table1_col2_row22
table1_col6_row23 = table1_col2_row23
table1_col6_row24 = table1_col2_row24

summ_col6 = (
    table1_col6_row4 + table1_col6_row5 + table1_col6_row6 + table1_col6_row7 + table1_col6_row8 +
    table1_col6_row9 + table1_col6_row10 + table1_col6_row11 + table1_col6_row12 + table1_col6_row13 +
    table1_col6_row14 + table1_col6_row15 + table1_col6_row16 + table1_col6_row17 + table1_col6_row18 +
    table1_col6_row19 + table1_col6_row20 + table1_col6_row21 + table1_col6_row22 + table1_col6_row23 +
    table1_col6_row24
)

table1_col6_row25 = summ_col6

col_6 = [
    f'2',
    f'Очищенный {in_data.name_comp1} из верха колонны {in_data.spec_equipment1}',
    f'6',
    'кг/ч',
    f'{table1_col6_row4:.4f}',
    f'{table1_col6_row5:.4f}',
    f'{table1_col6_row6:.4f}',
    f'{table1_col6_row7:.4f}',
    f'{table1_col6_row8:.4f}',
    f'{table1_col6_row9:.4f}',
    f'{table1_col6_row10:.4f}',
    f'{table1_col6_row11:.4f}',
    f'{table1_col6_row12:.4f}',
    f'{table1_col6_row13:.4f}',
    f'{table1_col6_row14:.4f}',
    f'{table1_col6_row15:.4f}',
    f'{table1_col6_row16:.4f}',
    f'{table1_col6_row17:.4f}',
    f'{table1_col6_row18:.4f}',
    f'{table1_col6_row19:.4f}',
    f'{table1_col6_row20:.4f}',
    f'{table1_col6_row21:.4f}',
    f'{table1_col6_row22:.4f}',
    f'{table1_col6_row23:.4f}',
    f'{table1_col6_row24:.4f}',
    f'{table1_col6_row25:.4f}'
]

table1_col7_row4 = table1_col6_row4 / table1_col6_row25 * 100
table1_col7_row5 = table1_col6_row5 / table1_col6_row25 * 100
table1_col7_row6 = table1_col6_row6 / table1_col6_row25 * 100
table1_col7_row7 = table1_col6_row7 / table1_col6_row25 * 100
table1_col7_row8 = table1_col6_row8 / table1_col6_row25 * 100
table1_col7_row9 = table1_col6_row9 / table1_col6_row25 * 100
table1_col7_row10 = table1_col6_row10 / table1_col6_row25 * 100
table1_col7_row11 = table1_col6_row11 / table1_col6_row25 * 100
table1_col7_row12 = table1_col6_row12 / table1_col6_row25 * 100
table1_col7_row13 = table1_col4_row13 / table1_col6_row25 * 100
table1_col7_row14 = table1_col4_row14 / table1_col6_row25 * 100
table1_col7_row15 = table1_col4_row15 / table1_col6_row25 * 100
table1_col7_row16 = table1_col4_row16 / table1_col6_row25 * 100
table1_col7_row17 = table1_col4_row17 / table1_col6_row25 * 100
table1_col7_row18 = table1_col4_row18 / table1_col6_row25 * 100
table1_col7_row19 = table1_col4_row19 / table1_col6_row25 * 100
table1_col7_row20 = table1_col4_row20 / table1_col6_row25 * 100
table1_col7_row21 = table1_col4_row21 / table1_col6_row25 * 100
table1_col7_row22 = table1_col4_row22 / table1_col6_row25 * 100
table1_col7_row23 = table1_col4_row23 / table1_col6_row25 * 100
table1_col7_row24 = table1_col4_row24 / table1_col6_row25 * 100

summ_col7 = (
    table1_col5_row4 + table1_col5_row5 + table1_col5_row6 + table1_col5_row7 + table1_col5_row8 +
    table1_col5_row9 + table1_col5_row10 + table1_col5_row11 + table1_col5_row12 + table1_col5_row13 +
    table1_col5_row14 + table1_col5_row15 + table1_col5_row16 + table1_col5_row17 + table1_col5_row18 +
    table1_col5_row19 + table1_col5_row20 + table1_col5_row21 + table1_col5_row22 + table1_col5_row23 +
    table1_col5_row24
)
table1_col7_row25 = summ_col7

col_7 = [
    None,
    None,
    '7',
    '%',
    f'{table1_col5_row4:.4f}',
    f'{table1_col5_row5:.4f}',
    f'{table1_col5_row6:.4f}',
    f'{table1_col5_row7:.4f}',
    f'{table1_col5_row8:.4f}',
    f'{table1_col5_row9:.4f}',
    f'{table1_col5_row10:.4f}',
    f'{table1_col5_row11:.4f}',
    f'{table1_col5_row12:.4f}',
    f'{table1_col5_row13:.4f}',
    f'{table1_col5_row14:.4f}',
    f'{table1_col5_row15:.4f}',
    f'{table1_col5_row16:.4f}',
    f'{table1_col5_row17:.4f}',
    f'{table1_col5_row18:.4f}',
    f'{table1_col5_row19:.4f}',
    f'{table1_col5_row20:.4f}',
    f'{table1_col5_row21:.4f}',
    f'{table1_col5_row22:.4f}',
    f'{table1_col5_row23:.4f}',
    f'{table1_col5_row24:.4f}',
    f'{table1_col5_row25:.4f}'
]


