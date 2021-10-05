import numpy as np
import matplotlib.pyplot as plt
import openpyxl
from collections import Counter

partii = openpyxl.open("Partii.xlsx",read_only=True)

#sheet 1 (6 partii)
sheet_1 = partii.worksheets[0]

def list_data(a, b):
    p_list = []
    for i in range(3, a):
        value = sheet_1.cell(row=i, column=b).value
        p_list.append(value)
    return p_list


def list_data_int(a, b):
    p_list = []
    for i in range(3, a):
        size = int(sheet_1.cell(row=i, column=b).value)
        p_list.append(size)
    return p_list


#p_100_s- списки размеров   p_100_std - сигма размеров p_100_cat - списки категорий
#партия 100
p_100_s = np.array(list_data_int(325, 1))
p_100_cat = list_data(325, 2)
p_100_std = np.std(p_100_s)
p_100_A = p_100_cat.count('A')
p_100_B = p_100_cat.count('B')
p_100_C = p_100_cat.count('C')
p_100_D = p_100_cat.count('D')
p_100_E = p_100_cat.count('E')

#партия 101
p_101_s = np.array(list_data_int(324, 6))
p_101_cat = list_data(324, 7)
p_101_std = np.std(p_101_s)
p_101_A = p_101_cat.count('A')
p_101_B = p_101_cat.count('B')
p_101_C = p_101_cat.count('C')
p_101_D = p_101_cat.count('D')
p_101_E = p_101_cat.count('E')

#партия 102
p_102_s = np.array(list_data_int(398, 11))
p_102_cat = list_data(398, 12)
p_102_std = np.std(p_102_s)
p_102_A = p_102_cat.count('A')
p_102_B = p_102_cat.count('B')
p_102_C = p_102_cat.count('C')
p_102_D = p_102_cat.count('D')
p_102_E = p_102_cat.count('E')

#партия 103
p_103_s = np.array(list_data_int(402, 16))
p_103_cat = list_data(402, 17)
p_103_std = np.std(p_103_s)
p_103_A = p_103_cat.count('A')
p_103_B = p_103_cat.count('B')
p_103_C = p_103_cat.count('C')
p_103_D = p_103_cat.count('D')
p_103_E = p_103_cat.count('E')

#партия 104
p_104_s = np.array(list_data_int(388, 21))
p_104_cat = list_data(388, 22)
p_104_std = np.std(p_104_s)
p_104_A = p_104_cat.count('A')
p_104_B = p_104_cat.count('B')
p_104_C = p_104_cat.count('C')
p_104_D = p_104_cat.count('D')
p_104_E = p_104_cat.count('E')

#партия 105
p_105_s = np.array(list_data_int(407, 26))
p_105_cat = list_data(407, 27)
p_105_std = np.std(p_105_s)
p_105_A = p_105_cat.count('A')
p_105_B = p_105_cat.count('B')
p_105_C = p_105_cat.count('C')
p_105_D = p_105_cat.count('D')
p_105_E = p_105_cat.count('E')

#партия 106
p_106_s = np.array(list_data_int(264, 31))
p_106_cat = list_data(264, 32)
p_106_std = np.std(p_106_s)
p_106_A = p_106_cat.count('A')
p_106_B = p_106_cat.count('B')
p_106_C = p_106_cat.count('C')
p_106_D = p_106_cat.count('D')
p_106_E = p_106_cat.count('E')

#общие списки
p_A_all = np.array([p_100_A, p_101_A, p_102_A, p_103_A, p_104_A, p_105_A, p_106_A])
p_B_all = np.array([p_100_B, p_101_B, p_102_B, p_103_B, p_104_B, p_105_B, p_106_B])
p_C_all = np.array([p_100_C, p_101_C, p_102_C, p_103_C, p_104_C, p_105_C, p_106_C])
p_D_all = np.array([p_100_D, p_101_D, p_102_D, p_103_D, p_104_D, p_105_D, p_106_D])
p_E_all = np.array([p_100_E, p_101_E, p_102_E, p_103_E, p_104_E, p_105_E, p_106_E])
p_std_all = np.array([p_100_std, p_101_std, p_102_std, p_103_std, p_104_std, p_105_std, p_106_std])


#second sheet
sheet_2 = partii.worksheets[1]

def list_data(a, b):
    p_list = []
    for i in range(3, a):
        value = sheet_2.cell(row=i, column=b).value
        p_list.append(value)
    return p_list


def list_data_int(a, b):
    p_list = []
    for i in range(3, a):
        size = int(sheet_2.cell(row=i, column=b).value)
        p_list.append(size)
    return p_list


#p_100_s- списки размеров   p_100_std - сигма размеров p_100_cat - списки категорий
#партия 107
p_107_s = np.array(list_data_int(389, 1))
p_107_cat = list_data(389, 2)
p_107_std = np.std(p_107_s)
p_107_A = p_107_cat.count('A')
p_107_B = p_107_cat.count('B')
p_107_C = p_107_cat.count('C')
p_107_D = p_107_cat.count('D')
p_107_E = p_107_cat.count('E')


#партия 108
p_108_s = np.array(list_data_int(652, 6))
p_108_cat = list_data(652, 7)
p_108_std = np.std(p_108_s)
p_108_A = p_108_cat.count('A')
p_108_B = p_108_cat.count('B')
p_108_C = p_108_cat.count('C')
p_108_D = p_108_cat.count('D')
p_108_E = p_108_cat.count('E')

#партия 109
p_109_s = np.array(list_data_int(505, 11))
p_109_cat = list_data(505, 12)
p_109_std = np.std(p_109_s)
p_109_A = p_109_cat.count('A')
p_109_B = p_109_cat.count('B')
p_109_C = p_109_cat.count('C')
p_109_D = p_109_cat.count('D')
p_109_E = p_109_cat.count('E')

p_A_all = np.append(p_A_all, [p_107_A, p_108_A, p_109_A])
p_B_all = np.append(p_B_all, [p_107_B, p_108_B, p_109_B])
p_C_all = np.append(p_C_all, [p_107_C, p_108_C, p_109_C])
p_D_all = np.append(p_D_all, [p_107_D, p_108_D, p_109_D])
p_E_all = np.append(p_E_all, [p_107_E, p_108_E, p_109_E])
p_std_all = np.append(p_std_all, [p_107_std, p_108_std, p_109_std])

print(p_A_all , p_B_all, p_C_all, p_D_all, p_E_all, p_std_all)
#все стандартные отклонения из списков
b = (np.mean(p_A_all * p_std_all) - np.mean(p_std_all) * np.mean(p_A_all)) / (np.mean (p_A_all ** 2) - np.mean(p_A_all) ** 2)
a = np.mean(p_std_all) - b * np.mean(p_A_all)
print(f'y ={a:.2f}+{b:.2f}x')
y = a + b * p_A_all

plt.scatter(p_A_all, p_std_all)
plt.plot(p_A_all, y)
plt.xlabel('Количество шкур А в партии')
plt.ylabel('Дисперсия размеров в партии')
plt.show
