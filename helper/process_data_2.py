import math
import copy
import pandas
import locale
from datetime import datetime
from helper.interpolator import *

locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')


def run(file_path):
    df = pandas.read_excel(file_path)
    df_clean = (df.iloc[10:41, 1:26]).dropna(how='all')

    detail = (df.iloc[3:7, 2:3])
    detail = detail.values.flatten()

    # Menghapus elemen 'Unnamed: 2' dan mengubah menjadi list
    detail = [x for x in detail]

    year = detail[2].split(' ')[2]

    # section1
    section1 = df_clean.values.tolist()

    dates = [datetime.strptime(row.pop(0), '%d %B %Y') for row in section1]

    # section2
    section2 = []
    for i, row in enumerate(section1):
        x1_positive = sum(row[6:18])
        x1_negative = sum(row[:6]) + sum(row[18:])
        y1_positive = sum(row[12:])
        y1_negative = sum(row[:12])
        x2_positive = sum(row[:3]) + sum(row[9:15]) + sum(row[21:])
        x2_negative = sum(row[3:9]) + sum(row[15:21])
        y2_positive = sum(row[:6]) + sum(row[12:18])
        y2_negative = sum(row[6:12]) + sum(row[18:24])
        x3_positive = row[0] + sum(row[5:7]) + sum(row[11:13]) + sum(row[17:19]) + row[23]
        x3_negative = sum(row[2:4]) + sum(row[8:10]) + sum(row[14:16]) + sum(row[20:22])
        y3_positive = sum(row[0:3]) + sum(row[6:9]) + sum(row[12:15]) + sum(row[18:21])
        y3_negative = sum(row[3:6]) + sum(row[9:12]) + sum(row[15:18]) + sum(row[21:])

        section2.append([x1_positive, x1_negative, y1_positive, y1_negative, x2_positive, 
                        x2_negative, y2_positive, y2_negative, x3_positive, x3_negative,
                        y3_positive, y3_negative])

    A10 = len(section2)
    A22 = len(section2)
    A27 = len(section2)
    A32 = len(section2)
    A37 = len(section2)
    A15 = len(section2)
    A35 = len(section2)

    # section3
    cx_1 = 2000
    cy_1 = 2000
    cx_2 = 2000
    cy_2 = 2000
    cx_4 = 500
    cy_4 = 500

    section3 = [
        [
            sum(row[0:2]),
            (row[0] - row[1] + cx_1) if A35 == 29 else None,
            (row[2] - row[3] + cy_1) if A35 == 29 else None,
            (row[4] - row[5] + cx_2) if A35 == 29 else None,
            (row[6] - row[7] + cy_2) if A35 == 29 else None,
            (row[8] - row[9] + cx_4) if A35 == 29 else None,
            (row[10] - row[11] + cy_4) if A35 == 29 else None
        ]
        for row in section2
    ]

    def measure_sec4(data, index, a, list_data_1, list_data_2=None):
        if a == 29:
            data_range = [data[row][index] for row in list_data_1]
            return sum(data_range)
        if a == 15 and list_data_2 is not None:
            data_range = [data[row][index] for row in list_data_2]
            return sum(data_range)
        return None

    dat_sec_4x_1 = measure_sec4(section3, 0, A10, range(0,29))
    dat_sec_4x_2 = measure_sec4(section3, 1, A10, range(0,29), range(7,27))
    dat_sec_4x_3 = A10 * cx_1
    dat_sec_4x_4 = measure_sec4(section3, 1, A10, [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28])
    dat_sec_4x_5 = measure_sec4(section3, 1, A10, [4,5,6,7,8,9,10,18,19,20,21,22,23,24], [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28])
    dat_sec_4x_6 = cx_1
    dat_sec_4x_7 = measure_sec4(section3, 1, A10, [8,9,10,11,12,13,22,23,24,25,26,27], [8,9,10,11,12,13])
    dat_sec_4x_8 = measure_sec4(section3, 1, A10, [1,2,3,4,5,6,15,16,17,18,19,20], [15,16,17,18,19,20])
    dat_sec_4x_9 = measure_sec4(section3, 1, A15, [2,3,4,5,6,12,13,14,15,16,22,23,24,25,26], [12,13,14,15,16])
    dat_sec_4x_10 = measure_sec4(section3, 1, A15, [0,1,7,8,9,10,11,17,18,19,20,21,27,28], [7,8,9,10,11,17,18,19,20,21])
    dat_sec_4x_11 = cx_1 if A15 == 29 else cx_1 * 5
    dat_sec_4x_12 = measure_sec4(section3, 1, A15, [0,1,2,3,4,10,11,12,13,19,20,21,22,23], [10,11,12,13,19,20,21])
    dat_sec_4x_13 = measure_sec4(section3, 1, A15, [5,6,7,8,9,15,16,17,18,24,25,26,27,28], [7,8,9,15,16,17,18])
    dat_sec_4x_14 = measure_sec4(section3, 3, A22, range(0,29), range(7,27))
    dat_sec_4x_15 = A22 * cx_2
    dat_sec_4x_16 = measure_sec4(section3, 3, A22, [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28], [11,12,13,14,15,16,17])
    dat_sec_4x_17 = measure_sec4(section3, 3, A22, [4,5,6,7,8,9,10,18,19,20,21,22,23,24], [7,8,9,10,18,19,20,21])
    dat_sec_4x_18 = cx_2
    dat_sec_4x_19 = measure_sec4(section3, 3, A27, [8,9,10,11,12,13,22,23,24,25,26,27],[8,9,10,11,12,13])
    dat_sec_4x_20 = measure_sec4(section3, 3, A27, [1,2,3,4,5,6,15,16,17,18,19,20],[15,16,17,18,19,20])
    dat_sec_4x_21 = measure_sec4(section3, 3, A27, [2,3,4,5,6,12,13,14,15,16,22,23,24,25,26],[12,13,14,15,16])
    dat_sec_4x_22 = measure_sec4(section3, 3, A27, [0,1,7,8,9,10,11,17,18,19,20,21,27,28],[7,8,9,10,11,17,18,19,20,21])
    dat_sec_4x_23 = cx_2 if A27 == 29 else cx_2 * 5
    dat_sec_4x_24 = measure_sec4(section3, 3, A27, [0,1,2,3,4,10,11,12,13,19,20,21,22,23],[10,11,12,13,19,20,21])
    dat_sec_4x_25 = measure_sec4(section3, 3, A27, [5,6,7,8,9,15,16,17,18,24,25,26,27,28],[7,8,9,15,16,17,18])
    dat_sec_4x_26 = measure_sec4(section3, 5, A32, [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28],[11,12,13,14,15,16,17])
    dat_sec_4x_27 = measure_sec4(section3, 5, A32, [4,5,6,7,8,9,10,18,19,20,21,22,23,24],[7,8,9,10,18,19,20,21])
    dat_sec_4x_28 = cx_4
    dat_sec_4x_29 = measure_sec4(section3, 5, A32, [8,9,10,11,12,13,22,23,24,25,26,27],[8,9,10,11,12,13])
    dat_sec_4x_30 = measure_sec4(section3, 5, A32, [1,2,3,4,5,6,15,16,17,18,19,20],[15,16,17,18,19,20,21])
    dat_sec_4x_31 = measure_sec4(section3, 5, A37, [0,1,5,6,7,8,13,14,15,20,21,22,23,27,28],[8,9,13,14,15,20,21])
    dat_sec_4x_32 = measure_sec4(section3, 5, A37, [2,3,4,9,10,11,12,16,17,18,19,24,25,26],[9,10,11,12,16,17,18,29])
    dat_sec_4x_33 = cx_4
    dat_sec_4x_34 = measure_sec4(section3, 5, A37, [4,5,6,11,12,13,18,19,20,25,26,27],[11,12,13,18,19,20])
    dat_sec_4x_35 = measure_sec4(section3, 5, A37, [1,2,3,8,9,10,15,16,17,22,23,24],[8,9,10,15,16,17])

    dat_sec_4y_1 = 0
    dat_sec_4y_2 = measure_sec4(section3, 2, A10, range(0,29), range(7,21))
    dat_sec_4y_3 = A10 * cy_1
    dat_sec_4y_4 = measure_sec4(section3, 2, A10, [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28],[11,12,13,14,15,16,17])
    dat_sec_4y_5 = measure_sec4(section3, 2, A10, [4,5,6,7,8,9,10,18,19,20,21,22,23,24],[7,8,9,10,18,19,20,21])
    dat_sec_4y_6 = cy_1
    dat_sec_4y_7 = measure_sec4(section3, 2, A10, [8,9,10,11,12,13,22,23,24,25,26,27], [8,9,10,11,12,13])
    dat_sec_4y_8 = measure_sec4(section3, 2, A10, [1,2,3,4,5,6,15,16,17,18,19,20], [15,16,17,18,19,20])
    dat_sec_4y_9 = measure_sec4(section3, 2, A15, [2,3,4,5,6,12,13,14,15,16,22,23,24,25,26], [12,13,14,15,16])
    dat_sec_4y_10 = measure_sec4(section3, 2, A15, [0,1,7,8,9,10,11,17,18,19,20,21,27,28], [7,8,9,10,11,17,18,19,20,21])
    dat_sec_4y_11 = cy_1 if A15 == 29 else cy_1 * 5
    dat_sec_4y_12 = measure_sec4(section3, 2, A15, [0,1,2,3,4,10,11,12,13,19,20,21,22,23], [10,11,12,13,19,20,21])
    dat_sec_4y_13 = measure_sec4(section3, 2, A15, [5,6,7,8,9,15,16,17,18,24,25,26,27,28], [7,8,9,15,16,17,18])
    dat_sec_4y_14 = measure_sec4(section3, 4, A22, range(0,29), range(7,27))
    dat_sec_4y_15 = A22 * cy_2
    dat_sec_4y_16 = measure_sec4(section3, 4, A22, [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28], [11,12,13,14,15,16,17])
    dat_sec_4y_17 = measure_sec4(section3, 4, A22, [4,5,6,7,8,9,10,18,19,20,21,22,23,24], [7,8,9,10,18,19,20,21])
    dat_sec_4y_18 = cy_2
    dat_sec_4y_19 = measure_sec4(section3, 4, A27, [8,9,10,11,12,13,22,23,24,25,26,27],[8,9,10,11,12,13])
    dat_sec_4y_20 = measure_sec4(section3, 4, A27, [1,2,3,4,5,6,15,16,17,18,19,20],[15,16,17,18,19,20])
    dat_sec_4y_21 = measure_sec4(section3, 4, A27, [2,3,4,5,6,12,13,14,15,16,22,23,24,25,26],[12,13,14,15,16])
    dat_sec_4y_22 = measure_sec4(section3, 4, A27, [0,1,7,8,9,10,11,17,18,19,20,21,27,28],[7,8,9,10,11,17,18,19,20,21])
    dat_sec_4y_23 = cy_2 if A27 == 29 else cy_2 * 5
    dat_sec_4y_24 = measure_sec4(section3, 4, A27, [0,1,2,3,4,10,11,12,13,19,20,21,22,23],[10,11,12,13,19,20,21])
    dat_sec_4y_25 = measure_sec4(section3, 4, A27, [5,6,7,8,9,15,16,17,18,24,25,26,27,28],[7,8,9,15,16,17,18])
    dat_sec_4y_26 = measure_sec4(section3, 6, A32, [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28],[11,12,13,14,15,16,17])
    dat_sec_4y_27 = measure_sec4(section3, 6, A32, [4,5,6,7,8,9,10,18,19,20,21,22,23,24],[7,8,9,10,18,19,20,21])
    dat_sec_4y_28 = cy_4
    dat_sec_4y_29 = measure_sec4(section3, 6, A32, [8,9,10,11,12,13,22,23,24,25,26,27],[8,9,10,11,12,13])
    dat_sec_4y_30 = measure_sec4(section3, 6, A32, [1,2,3,4,5,6,15,16,17,18,19,20],[15,16,17,18,19,20,21])
    dat_sec_4y_31 = measure_sec4(section3, 6, A37, [0,1,5,6,7,8,13,14,15,20,21,22,23,27,28],[8,9,13,14,15,20,21])
    dat_sec_4y_32 = measure_sec4(section3, 6, A37, [2,3,4,9,10,11,12,16,17,18,19,24,25,26],[9,10,11,12,16,17,18,29])
    dat_sec_4y_33 = cy_4
    dat_sec_4y_34 = measure_sec4(section3, 6, A37, [4,5,6,11,12,13,18,19,20,25,26,27],[11,12,13,18,19,20])
    dat_sec_4y_35 = measure_sec4(section3, 6, A37, [1,2,3,8,9,10,15,16,17,22,23,24],[8,9,10,15,16,17])

    # section_3 = [
    #     ["00", "+", 100967.0548, None, 100967.0548, None],
    #     ["10", "+", 57040.67291, 67498.80469, -959.3270898, None],
    #     [None, "-", 58000, 58000, None, 9498.804693],
    #     ["12", "+", 30187.91216, 36546.64474, None, 3594.484792],
    #     [None, "-", 26852.76075, 30952.15995, 1335.151419, None],
    #     ["29", "(-)  (+)", 2000, 2000, None, None],
    #     ["1b", "+", 25528.86764, 26595.07553, 3898.173086, -2539.649243],
    #     [None, "-", 21630.69455, 29134.72477, None, None],
    #     ["13", "+", 30047.33663, 35248.7482, None, 998.6916995],
    #     [None, "-", 26993.33628, 32250.0565, 1054.00035, None],
    #     ["29", "(-)  (+)", 2000, 2000, None, None],
    #     ["1c", "+", 27399.32573, 32092.19264, -171.4695604, -788.6917405],
    #     [None, "-", 27570.79529, 32880.88438, None, None],
    #     ["20", "+", 69665.062, 53268.5847, 11665.062, None],
    #     [None, "-", 58000, 58000, None, -4731.415296],
    #     ["22", "+", 31723.78444, 29089.64042, -8217.493126, 2910.696129],
    #     [None, "-", 37941.27756, 24178.94429, None, None],
    #     ["29", "(-)  (+)", 2000, 2000, None, None],
    #     ["2b", "+", 30657.81599, 26485.67715, 3214.905949, None],
    #     [None, "-", 27442.91004, 17504.7914, None, 8980.885749],
    #     ["23", "+", 35416.74226, 28469.01076, None, 1669.436813],
    #     [None, "-", 34248.31974, 24799.57395, -831.5774784, None],
    #     ["29", "(-)  (+)", 2000, 2000, None, None],
    #     ["2c", "+", 34416.00954, 25305.13092, 1096.311023, -535.7400476],
    #     [None, "-", 33319.69851, 25840.87097, None, None],
    #     ["42", "+", 7384.865683, 7557.612331, -264.643106, None],
    #     [None, "-", 7149.508789, 6985.676678, None, 71.93565312],
    #     ["29", "(-)  (+)", 500, 500, None, None],
    #     ["4b", "+", 6034.024863, 6363.643023, None, None],
    #     [None, "-", 5987.403657, 5667.918853, 46.62120616, 695.7241696],
    #     ["44", "+", 7580.370723, 7549.966083, None, 56.64315705],
    #     [None, "-", 6954.003748, 6993.322926, 126.3669747, None],
    #     ["29", "(-)  (+)", 500, 500, None, None],
    #     ["4d", "+", 6017.330351, 5981.370501, 13.23218204, None],
    #     [None, "-", 6004.098169, 6050.191375, None, -68.82087351]
    # ]

    section_3 = [
        ["00", "+", dat_sec_4x_1, None, dat_sec_4x_1, None],
        ["10", "+", dat_sec_4x_2, dat_sec_4y_2, (dat_sec_4x_2 - dat_sec_4x_3) if A10 == 29 else 0, None],
        [None, "-", dat_sec_4x_3, dat_sec_4y_3, None, dat_sec_4y_2 - dat_sec_4y_3],
        ["12", "+", dat_sec_4x_4, dat_sec_4y_4, None, (dat_sec_4y_4 - dat_sec_4y_5 - dat_sec_4y_6) if A10 == 29 else 0],
        [None, "-", dat_sec_4x_5, dat_sec_4y_5, (dat_sec_4x_4 - dat_sec_4x_5 - dat_sec_4x_6) if A10 == 29 else 0, None],
        ["29", "(-)  (+)", dat_sec_4x_6, dat_sec_4y_6, None, None],
        ["1b", "+", dat_sec_4x_7, dat_sec_4y_7, (dat_sec_4x_7 - dat_sec_4x_8) if A10 == 29 else 0, (dat_sec_4y_7 - dat_sec_4y_8) if A10 == 29 else 0],
        [None, "-", dat_sec_4x_8, dat_sec_4y_8, None, None],
        ["13", "+", dat_sec_4x_9, dat_sec_4y_9, None, (dat_sec_4y_9 - dat_sec_4y_10 - dat_sec_4y_11) if A15 == 29 else 0],
        [None, "-", dat_sec_4x_10, dat_sec_4y_10, (dat_sec_4x_9 - dat_sec_4x_10 - dat_sec_4x_11) if A15 == 29 else 0, None],
        ["29", "(-)  (+)", dat_sec_4x_11, dat_sec_4y_11, None, None],
        ["1c", "+", dat_sec_4x_12, dat_sec_4y_12, (dat_sec_4x_12 - dat_sec_4x_13) if A15 == 29 else 0, (dat_sec_4y_12 - dat_sec_4y_13) if A15 == 29 else 0],
        [None, "-", dat_sec_4x_13, dat_sec_4y_13, None, None],
        ["20", "+", dat_sec_4x_14, dat_sec_4y_14, dat_sec_4x_14-dat_sec_4x_15, None],
        [None, "-", dat_sec_4x_15, dat_sec_4y_15, None, (dat_sec_4y_14 - dat_sec_4x_15) if A22 == 29 else 0],
        ["22", "+", dat_sec_4x_16, dat_sec_4y_16, (dat_sec_4x_16-dat_sec_4x_17-dat_sec_4x_18) if A22 == 29 else 0, (dat_sec_4y_16-dat_sec_4y_17-dat_sec_4y_18) if A22 == 29 else 0],
        [None, "-", dat_sec_4x_17, dat_sec_4y_17, None, None],
        ["29", "(-)  (+)", dat_sec_4x_18, dat_sec_4y_18, None, None],
        ["2b", "+", dat_sec_4x_19, dat_sec_4y_19, (dat_sec_4x_19-dat_sec_4x_20) if A27 == 29 else 0, None],
        [None, "-", dat_sec_4x_20, dat_sec_4y_20, None, (dat_sec_4y_19-dat_sec_4y_20) if A27 == 29 else 0],
        ["23", "+", dat_sec_4x_21, dat_sec_4y_21, None, (dat_sec_4y_21-dat_sec_4y_22-dat_sec_4y_23) if A27 == 29 else 0],
        [None, "-", dat_sec_4x_22, dat_sec_4y_22, (dat_sec_4x_21-dat_sec_4x_22-dat_sec_4x_23) if A27 == 29 else (dat_sec_4x_21-dat_sec_4x_22+dat_sec_4x_23), None],
        ["29", "(-)  (+)", dat_sec_4x_23, dat_sec_4y_23, None, None],
        ["2c", "+", dat_sec_4x_24, dat_sec_4y_24, (dat_sec_4x_24-dat_sec_4x_25) if A27 == 29 else 0, dat_sec_4y_24-dat_sec_4y_25],
        [None, "-", dat_sec_4x_25, dat_sec_4y_25, None, None],
        ["42", "+", dat_sec_4x_26, dat_sec_4y_26, (dat_sec_4x_26-dat_sec_4x_27-dat_sec_4x_28) if A32 == 29 else 0, None],
        [None, "-", dat_sec_4x_27, dat_sec_4y_27, None, (dat_sec_4y_26-dat_sec_4y_27-dat_sec_4y_28) if A32 == 29 else (dat_sec_4y_26-dat_sec_4y_27+dat_sec_4y_28)],
        ["29", "(-)  (+)", dat_sec_4x_28, dat_sec_4y_28, None, None],
        ["4b", "+", dat_sec_4x_29, dat_sec_4y_29, None, None],
        [None, "-", dat_sec_4x_30, dat_sec_4y_30, (dat_sec_4x_29-dat_sec_4x_30) if A32 == 29 else 0, (dat_sec_4y_29-dat_sec_4y_30) if A32 == 29 else 0],
        ["44", "+", dat_sec_4x_31, dat_sec_4y_31, None, (dat_sec_4y_31-dat_sec_4y_32-dat_sec_4y_33) if A37 == 29 else 0],
        [None, "-", dat_sec_4x_32, dat_sec_4y_32, (dat_sec_4x_31-dat_sec_4x_32-dat_sec_4x_33) if A37 == 29 else 0, None],
        ["29", "(-)  (+)", dat_sec_4x_33, dat_sec_4y_33, None, None],
        ["4d", "+", dat_sec_4x_34, dat_sec_4y_34, dat_sec_4x_34-dat_sec_4x_35, None],
        [None, "-", dat_sec_4x_35, dat_sec_4y_35, None, dat_sec_4y_34-dat_sec_4y_35]
    ]

    # for row in section_3:
    #     print(row)

    table_V_VI = [
        section_3[0][4],
        section_3[1][4] if A10 == 29 else section_3[2][4],
        (section_3[4][4] - section_3[6][5]) if A10 == 29 else (section_3[3][4] - section_3[7][5] if A10 == 15 else None),
        (section_3[8][2] - section_3[9][2] + section_3[10][2]) - (section_3[11][3] - section_3[12][3]),
        section_3[13][2] - section_3[14][2],
        (section_3[15][4] - section_3[19][5]) if A10 == 29 else (section_3[16][4] - section_3[18][5] if A10 == 15 else None),
        section_3[21][4] - section_3[23][5],
        ((section_3[25][2] - section_3[26][2] - section_3[27][2]) - (section_3[28][3] - section_3[29][3])) if A10 == 29 else (((section_3[25][2] - section_3[26][2] + section_3[27][2]) - (section_3[28][3] - section_3[29][3])) if A10 == 15 else None),
        (section_3[31][4] - section_3[34][5]) if A10 == 29 else (section_3[30][4] - section_3[32][5] if A10 == 15 else None),
        section_3[2][5] if A10 == 29 else section_3[1][5],
        (section_3[3][5] + section_3[6][4]) if A10 == 29 else (section_3[4][5] + section_3[7][4] if A10 == 15 else None),
        (section_3[8][5] + section_3[11][4]) if A15 == 29 else (section_3[9][5] + section_3[12][4] if A15 == 15 else None),
        section_3[14][5]  if A22 == 29 else section_3[13][5] ,
        (section_3[15][5] + section_3[18][4]) if A22 == 29 else (section_3[16][5] + section_3[19][4] if A22 == 15 else None),
        (section_3[20][5] + section_3[23][4]) if A27 == 29 else (section_3[21][5] + section_3[23][4] if A27 == 15 else None),
        (section_3[26][5] + section_3[29][4]) if A32 == 29 else (section_3[26][5] + section_3[28][4] if A32 == 15 else None),
        (section_3[30][5] + section_3[33][4]) if A37 == 29 else (section_3[31][5] + section_3[33][4] if A37 == 15 else None)
    ]

    result1 = table_V_VI[2] * 0.07 if A10 == 29 else (table_V_VI[2] * 0.09 if A10 == 15 else 0)
    result2 = table_V_VI[2] * -0.02 if A10 == 29 else (table_V_VI[2] * -0.09 if A10 == 15 else 0)
    result3 = table_V_VI[2] * 1 if A10 == 29 else (table_V_VI[2] * 1 if A10 == 15 else 0)
    result4 = table_V_VI[2] * 0.02 if A10 == 29 else (table_V_VI[2] * 0.02 if A10 == 15 else 0)
    result5 = table_V_VI[4] * -0.03 if A10 == 29 else (table_V_VI[4] * -0.15 if A10 == 15 else None)
    result6 = table_V_VI[4] * 1 if A10 == 29 else (table_V_VI[4] * 1 if A10 == 15 else None)
    result7 = table_V_VI[4] * -0.03 if A10 == 29 else (table_V_VI[4] * 0.29 if A10 == 15 else None)
    result8 = table_V_VI[5] * 1
    result9 = table_V_VI[5] * 0.015 if A10 == 29 else (table_V_VI[5] * -0.14 if A10 == 15 else None)
    result10 = table_V_VI[5] * 0.038 if A10 == 29 else (table_V_VI[5] * -0.61 if A10 == 15 else None)
    result11 = table_V_VI[5] * 0.002 if A10 == 29 else (table_V_VI[5] * -0.02 if A10 == 15 else None)
    result12 = table_V_VI[5] * -0.058 if A10 == 29 else (table_V_VI[5] * -0.03 if A10 == 15 else None)
    result13 = table_V_VI[5] * -0.035 if A10 == 29 else (table_V_VI[5] * -0.03 if A10 == 15 else None)
    result14 = table_V_VI[6] * -0.06 if A10 == 29 else (table_V_VI[6] * -0.65 if A10 == 15 else None)
    result15 = table_V_VI[6] * 1 if A10 == 29 else (table_V_VI[6] * 1 if A10 == 15 else None)
    result16 = table_V_VI[7] * 0.03 if A10 == 29 else (table_V_VI[7] * 0.01 if A10 == 15 else None)
    result17 = table_V_VI[7] * 1
    result18 = table_V_VI[8] * 1 if A10 == 29 else (table_V_VI[8] * 1.01 if A10 == 15 else None)
    result19 = table_V_VI[8] * 0.08 if A10 == 29 else (table_V_VI[8] * -0.05 if A10 == 15 else None)

    section5 = [
        [table_V_VI[0], 0, 0, 0, 0, 0, 0, 0],							
		[0, 0, 0, 0, table_V_VI[1] * 1 if A10 in [29, 15] else 0, table_V_VI[1] * (-0.08) if A10 == 29 else (table_V_VI[1] * (-0.07) if A10 == 15 else 0), 0, 0],
	    [0, result1, 0, 0, result2, result3, 0, result4],
		[0, 0, 0, 0, 0, 0, 0, 0],		
	    [0, round(result5), result6, round(result7), 0, 0, 0, 0],				
	    [0, result8, result9, result10, result11, result12, 0, result13],
	    [0, round(result14, 1), 0, result15, 0, 0, 0, 0],
	    [0, round(result16, 1), 0, 0, 0, 0, 0, result17],
	    [0, 0, 0, 0, 0, 0, result18, result19],
    ]

    result20 = table_V_VI[9] * 1 if A10 == 29 else (table_V_VI[9] * 1.01 if A10 == 15 else None)
    result21 = table_V_VI[9] * -0.08
    result22 = table_V_VI[10] * 0.07 if A10 == 29 else (table_V_VI[11] * 0.05 if A10 == 15 else None)
    result23 = table_V_VI[10] * -0.02 if A10 == 29 else (table_V_VI[11] * -0.12 if A10 == 15 else None)
    result24 = table_V_VI[10] * 1 if A10 == 29 else (table_V_VI[11] * 1.05 if A10 == 15 else None)
    result25 = table_V_VI[10] * 0.03 if A10 == 29 else (table_V_VI[11] * 0.01 if A10 == 15 else None)
    result26 = table_V_VI[12] * -0.03 if A10 == 29 else (table_V_VI[12] * -0.16 if A10 == 15 else None)
    result27 = table_V_VI[12] * 1
    result28 = table_V_VI[12] * -0.03 if A10 == 29 else (table_V_VI[12] * 0.3 if A10 == 15 else None)
    result29 = table_V_VI[13] * 1 if A10 == 29 else (table_V_VI[13] * 1.04 if A10 == 15 else None)
    result30 = table_V_VI[13] * 0.015 if A10 == 29 else (table_V_VI[13] * -0.15 if A10 == 15 else None)
    result31 = table_V_VI[13] * 0.032 if A10 == 29 else (table_V_VI[13] * -0.64 if A10 == 15 else None)
    result32 = table_V_VI[13] * -0.057 if A10 == 29 else (table_V_VI[13] * -0.1 if A10 == 15 else None)
    result33 = table_V_VI[13] * -0.035 if A10 == 29 else (table_V_VI[13] * -0.02 if A10 == 15 else None)
    result34 = table_V_VI[14] * -0.06 if A10 == 29 else (table_V_VI[14] * -0.7 if A10 == 15 else None)
    result35 = table_V_VI[14] * 1 if A10 == 29 else (table_V_VI[14] * 1.03 if A10 == 15 else None)
    result36 = table_V_VI[15] * 0.03 if A10 == 29 else (table_V_VI[15] * 0.02 if A10 == 15 else None)
    result37 = table_V_VI[15] * 0.01 if A10 == 29 else (table_V_VI[15] * 0.11 if A10 == 15 else None)
    result38 = table_V_VI[15] * 1 if A10 == 29 else (table_V_VI[15] * 1 if A10 == 15 else None)
    result39 = table_V_VI[16] * 1 if A10 == 29 else (table_V_VI[16] * 1 if A10 == 15 else None)
    result40 = table_V_VI[16] * 0.08 if A10 == 29 else (table_V_VI[16] * -0.06 if A10 == 15 else None)

    section6 = [						
		[0, 0, 0, 0, round(result20, 2),round(result21,2), 0, 0],
	    [0, result22, 0, 0, result23, result24, 0, result25],
		[0, 0, 0, 0, 0, 0, 0, 0],		
	    [0, result26, result27, result28, 0, 0, 0, 0],				
	    [0, result29, result30, result31, 0, result32, 0, result33],
	    [0, result34, 0,result35, 0, 0, 0, 0],
	    [0, result36, 0, 0, 0, 0, result37, result38],
	    [0, 0, 0, 0, 0, 0, result39, result40],
    ]

    sum_PR = [
        [0, 0, 0, 0, 0, 0, 0, 0],
        [0, 0, 0, 0, 0, 0, 0, 0]
    ]

    for i in range(len(sum_PR)):
        for j in range(len(sum_PR[i])):
            if i == 0:
                for k in range(len(section5)):
                    sum_PR[i][j] += section5[k][j]
            if i == 1:
                for k in range(len(section6)):
                    sum_PR[i][j] += section6[k][j]

    distance_PR = []

    for i in range(len(sum_PR[0])):
        x1_squared = sum_PR[1][i] * sum_PR[1][i]
        x0_squared = sum_PR[0][i] * sum_PR[0][i]
        distance = math.sqrt(abs(x1_squared + x0_squared))
        distance_PR.append(distance)

        kabisat_constant = [
            [1, 31, 31, 15, 9, 2009],
            [2, 59, 60, False, False],
            [3, 90, 91, 257, 258],
            [4, 120, 121, 257, 258, 257],
            [5, 151, 152],
            [6, 181, 182, 257],
            [7, 212, 213, 257, 257],
            [8, 243, 244, 257],
            [9, 273, 274, 257],
            [10, 304, 305],
            [11, 334, 335],
            [12, 365, 366]
        ]

    dth  = [257, 2009, 27, 27]
    const_s = 277.025+(129.38481*(dth[1]-1900))+(13.1764*(dth[0]+dth[-1]))
    const_h = 280.19-(0.23872*(dth[1]-1900))+(0.98565*(dth[0]+dth[-1]))
    const_p = 334.385+(40.66249*(dth[1]-1900))+(0.1114*(dth[0]+dth[-1]))
    const_n = 259.157-(19.32818*(dth[1]-1900))-(0.05295*(dth[0]+dth[-1]))

    const_V = []
    const_u = []
    const_Vreal = []
    const_f = []
    const_V.append(-2*const_s+2*const_h)
    const_V.append(0)
    const_V.append(-3*const_s+2*const_h+const_p)
    const_V.append(2*const_h)
    const_V.append(const_h+90)
    const_V.append(-2*const_s+const_h+270)
    const_V.append(270-const_h)
    const_V.append(2*const_V[0])
    const_V.append(const_V[0])

    R42 = const_n
    const_u.append(-2.14 * math.sin(R42 * math.pi / 180))
    const_u.append(0)
    const_u.append(const_u[0])
    const_u.append((-17.74 * math.sin(R42 * math.pi / 180)) + (0.68 * math.sin(2 * R42 * math.pi / 180)) - (0.04 * math.sin(3 * R42 * math.pi / 180)))
    const_u.append((-8.86 * math.sin(R42 * math.pi / 180)) + (0.68 * math.sin(2 * R42 * math.pi / 180)) - (0.07 * math.sin(3 * R42 * math.pi / 180)))
    const_u.append((10.8 * math.sin(R42 * math.pi / 180)) - (1.34 * math.sin(2 * R42 * math.pi / 180)) + (0.19 * math.sin(3 * R42 * math.pi / 180)))
    const_u.append(0)
    const_u.append(2*const_u[0])
    const_u.append(const_u[0])

    const_Vreal = [N48 - (math.floor(N48 / 360) * 360) for N48 in const_V]

    const_f.append(1.0004 - (0.0373 * math.cos(R42 * math.pi / 180)) + (0.0002 * math.cos(2 * R42 * math.pi / 180)))
    const_f.append(1)
    const_f.append(const_f[0])
    const_f.append(1.0241 + (0.2863 * math.cos(R42 * math.pi / 180)) + (0.0083 * math.cos(2 * R42 * math.pi / 180)) - (0.0015 * math.cos(3 * R42 * math.pi / 180)))
    const_f.append(1.006 + (0.115 * math.cos(R42 * math.pi / 180)) - (0.0088 * math.cos(2 * R42 * math.pi / 180)) + (0.0006 * math.cos(3 * R42 * math.pi / 180)))
    const_f.append(1.0089 + (0.1871 * math.cos(R42 * math.pi / 180)) - (0.0147 * math.cos(2 * R42 * math.pi / 180)) + (0.0014 * math.cos(3 * R42 * math.pi / 180)))
    const_f.append(1)
    const_f.append(const_f[0]*const_f[0])
    const_f.append(const_f[0])

    print('const_V', [round(x,3) for x in const_V])
    print('const_u', [round(x,3) for x in const_u])
    print('const_Vreal', [round(x,3) for x in const_Vreal])
    print('const_f', [round(x,3) for x in const_f])

    A27 = 29
    if A27 == 29:
        new_P = [696,559,448,566,439,565,507,535]
        new_f = [0, const_f[0], const_f[1], const_f[2], const_f[4], const_f[5], const_f[7], const_f[8]]
        new_u = [0, const_u[0], const_u[1], const_u[2], const_u[4], const_u[5], const_u[7], const_u[8]]
        new_v = [0, const_Vreal[0], const_Vreal[1], const_Vreal[2], const_Vreal[4], const_Vreal[5], const_Vreal[7], const_Vreal[8]]
    else:
        new_P = [360,175,214,166,217,177,273,280]
        # edit
        new_f = [0, const_f[0], const_f[1], const_f[2], const_f[4], const_f[5], const_f[7], const_f[8]]
        new_u = [0, const_u[0], const_u[1], const_u[2], const_u[4], const_u[5], const_u[7], const_u[8]]
        new_v = [0, const_Vreal[0], const_Vreal[1], const_Vreal[2], const_Vreal[4], const_Vreal[5], const_Vreal[7], const_Vreal[8]]

    new_p = [0,333,345,327,173,160,307,318]
    tan_PR = []
    # rums = math.atan2(2, 3) * 180 / math.pi

    rums = []

    # Iterasi sebanyak len(sum_PR[0])
    for i in range(len(sum_PR[0])):
        angle = (math.atan2(sum_PR[0][i], sum_PR[1][i]) * 180 / math.pi) - 90
        rums.append(angle)

    rums = [-x for x in rums]
    new_r = [x - (math.floor(x / 360) * 360) for x in rums]
    # new_g360 = [x - (math.floor(x / 360) * 360) for x in rums]

    const_w_1 = [
        [0, 0.7, -0.214, 0, 0, 0.331, 0, 0, 1.184],
        [10, -6.6, -0.192, 10, -2.5, 0.327, 10, 1, 1.182],
        [20, -12.3, -0.131, 20, -4.9, 0.316, 20, 1.6, 1.174],
        [30, -15.5, -0.046, 30, -7.3, 0.297, 30, 4.6, 1.163],
        [40, -16.5, 0.047, 40, -9.6, 0.271, 40, 5.9, 1.147],
        [50, -15.6, 0.134, 50, -11.8, 0.239, 50, 7.2, 1.127],
        [60, -13.4, 0.207, 60, -13.8, 0.201, 60, 8.3, 1.104],
        [70, -10.3, 0.258, 70, -15.6, 0.157, 70, 9.2, 1.077],
        [80, -6.6, 0.284, 80, -17.1, 0.107, 80, 9.9, 1.048],
        [90, -2.6, 0.284, 90, -18.3, 0.053, 90, 10.4, 1.017],
        [100, 1.6, 0.256, 100, -19.1, -0.003, 100, 10.6, 0.984],
        [110, 5.6, 0.204, 110, -19.3, -0.06, 110, 10.4, 0.953],
        [120, 9.2, 0.131, 120, -19, -0.118, 120, 10, 0.922],
        [130, 12, 0.041, 130, -17.8, -0.173, 130, 9.1, 0.893],
        [140, 13.7, -0.058, 140, -15.9, -0.224, 140, 7.8, 0.867],
        [150, 13.6, -0.157, 150, -13.1, -0.268, 150, 6.2, 0.846],
        [160, 11.2, -0.245, 160, -9.3, -0.302, 160, 4.3, 0.83],
        [170, 6, -0.307, 170, -4.9, -0.323, 170, 2.2, 0.819],
        [180, -0.9, -0.33, 180, 0, -0.331, 180, 0, 0.816],
        [190, -7.8, -0.308, 190, 4.9, -0.323, 190, -2.2, 0.819],
        [200, -12.6, -0.247, 200, 9.3, -0.302, 200, -4.3, 0.83],
        [210, -14.9, -0.163, 210, 13.1, -0.268, 210, -6.2, 0.846],
        [220, -14.8, -0.067, 220, 15.9, -0.224, 220, -7.8, 0.867],
        [230, -13, 0.029, 230, 17.8, -0.173, 230, -9.1, 0.893],
        [240, -9.8, 0.115, 240, 19, -0.118, 240, -10, 0.922],
        [250, -6, 0.186, 250, 19.3, -0.06, 250, -10.4, 0.953],
        [260, -1.8, 0.236, 260, 19.1, -0.003, 260, -10.6, 0.984],
        [270, 2.6, 0.263, 270, 18.3, 0.053, 270, -10.4, 1.017],
        [280, 6.9, 0.265, 280, 17.1, 0.107, 280, -9.9, 1.048],
        [290, 10.8, 0.241, 290, 15.6, 0.157, 290, -9.2, 1.077],
        [300, 14.1, 0.192, 300, 13.8, 0.201, 300, -8.3, 1.104],
        [310, 16.5, 0.124, 310, 11.8, 0.239, 310, -7.2, 1.127],
        [320, 17.5, 0.039, 320, 9.6, 0.271, 320, -5.9, 1.147],
        [330, 16.8, -0.051, 330, 7.3, 0.297, 330, -4.6, 1.163],
        [340, 13.7, -0.133, 340, 4.9, 0.316, 340, -3.1, 1.174],
        [350, 8, -0.193, 350, 2.5, 0.327, 350, -1.6, 1.182],
        [360, 0.7, -0.214, 360, 0, 0.331, 360, 0, 1.184],
    ]

    # S0, MS4
    V_s0ms4 = copy.deepcopy(new_v[4])
    u_s0ms4 = copy.deepcopy(new_u[4])
    V_plus_u_s0ms4 = V_s0ms4 + u_s0ms4

    wf_s0ms4 = [interpolate_0(V_plus_u_s0ms4, const_w_1, 1, 0, 60), interpolate_0(V_plus_u_s0ms4, const_w_1, 1, 60, 120), 
                interpolate_0(V_plus_u_s0ms4, const_w_1, 1, 120, 180), interpolate_0(V_plus_u_s0ms4, const_w_1, 1, 180, 240), 
                interpolate_0(V_plus_u_s0ms4, const_w_1, 1, 240, 300), interpolate_0(V_plus_u_s0ms4, const_w_1, 1, 300, 360)]
    wf_s0ms4 = [x for x in wf_s0ms4 if x is not None]
    wf_s0ms4 = wf_s0ms4[0]

    Wf_s0ms4 = [interpolate_0(V_plus_u_s0ms4, const_w_1, 2, 0, 60), interpolate_0(V_plus_u_s0ms4, const_w_1, 2, 60, 120), 
                interpolate_0(V_plus_u_s0ms4, const_w_1, 2, 120, 180), interpolate_0(V_plus_u_s0ms4, const_w_1, 2, 180, 240), 
                interpolate_0(V_plus_u_s0ms4, const_w_1, 2, 240, 300), interpolate_0(V_plus_u_s0ms4, const_w_1, 2, 300, 360)]
    Wf_s0ms4 = [x for x in Wf_s0ms4 if x is not None]
    Wf_s0ms4 = Wf_s0ms4[0]

    f_s0ms4 = const_f[3]
    w_s0ms4 = wf_s0ms4 * f_s0ms4
    W_s0ms4 = Wf_s0ms4 * f_s0ms4
    W_s0ms4_plus_1 = Wf_s0ms4 + 1
    # print(V_s0ms4, u_s0ms4, V_plus_u_s0ms4, wf_s0ms4, Wf_s0ms4, f_s0ms4,w_s0ms4, W_s0ms4, W_s0ms4_plus_1)

    # K1
    v2_k1 = new_v[4] * 2
    u_k1 = new_u[4]
    v2_plus_u_k1 = (v2_k1 + u_k1) - (int((v2_k1 + u_k1) / 360) * 360)
    wf_k1 = [interpolate_1(v2_plus_u_k1, const_w_1, 4, 0, 60), interpolate_1(v2_plus_u_k1, const_w_1, 4, 60, 120), 
              interpolate_1(v2_plus_u_k1, const_w_1, 4, 120, 180), interpolate_1(v2_plus_u_k1, const_w_1, 4, 180, 240), 
              interpolate_1(v2_plus_u_k1, const_w_1, 4, 240, 300), interpolate_1(v2_plus_u_k1, const_w_1, 4, 300, 360)]
    wf_k1 = [x for x in wf_k1 if x is not None]
    wf_k1 = wf_k1[0]

    Wf_k1 = [interpolate_1(v2_plus_u_k1, const_w_1, 5, 0, 60), interpolate_1(v2_plus_u_k1, const_w_1, 5, 60, 120), 
              interpolate_1(v2_plus_u_k1, const_w_1, 5, 120, 180), interpolate_1(v2_plus_u_k1, const_w_1, 5, 180, 240), 
              interpolate_1(v2_plus_u_k1, const_w_1, 5, 240, 300), interpolate_1(v2_plus_u_k1, const_w_1, 5, 300, 360)]
    Wf_k1 = [x for x in Wf_k1 if x is not None]
    Wf_k1 = Wf_k1[0]

    f_k1 = new_f[4]
    w_k1 = wf_k1 / f_k1
    W_k1 = Wf_k1 / f_k1
    W_k1_plus_1 = W_k1 + 1
    # print(v2_k1, u_k1, v2_plus_u_k1, wf_k1, Wf_k1, f_k1,w_k1, W_k1, W_k1_plus_1)

    # N2
    v3_n2 = new_v[1] * 3
    v2_n2 = new_v[3] * 2
    v3_minus_v2 = (abs(v3_n2 - v2_n2)) - (int((abs(v3_n2 - v2_n2)) / 360) * 360)

    w_n2 = [interpolate_2(v3_minus_v2, const_w_1, 7, 0, 60), interpolate_2(v3_minus_v2, const_w_1, 7, 60, 120), 
              interpolate_2(v3_minus_v2, const_w_1, 7, 120, 180), interpolate_2(v3_minus_v2, const_w_1, 7, 180, 240), 
              interpolate_2(v3_minus_v2, const_w_1, 7, 240, 300), interpolate_2(v3_minus_v2, const_w_1, 7, 300, 360)]
    w_n2 = [x for x in w_n2 if x is not None]
    w_n2 = w_n2[0]
    w_n2_plus_1 = [interpolate_2(v3_minus_v2, const_w_1, 8, 0, 60), interpolate_2(v3_minus_v2, const_w_1, 8, 60, 120), 
              interpolate_2(v3_minus_v2, const_w_1, 8, 120, 180), interpolate_2(v3_minus_v2, const_w_1, 8, 180, 240), 
              interpolate_2(v3_minus_v2, const_w_1, 8, 240, 300), interpolate_2(v3_minus_v2, const_w_1, 8, 300, 360)]
    w_n2_plus_1 = [x for x in w_n2_plus_1 if x is not None]
    w_n2_plus_1 = w_n2_plus_1[0]

    # print(v3_n2, v2_n2, v3_minus_v2, w_n2, w_n2_plus_1)

    new_w = [0, 0, w_s0ms4, w_n2, w_k1, 0, 0, w_s0ms4]
    new_w_plus_1 = [0, 1, W_s0ms4_plus_1, w_n2_plus_1, W_k1_plus_1, 1, 1, W_s0ms4_plus_1]

    A = []
    g = []
    for i in range(len(new_w)):
        g.append(new_v[i] + new_u[i] + new_w[i] + new_p[i] + new_r[i])

        if i == 0:
            A.append(distance_PR[i]/new_P[i])
        else:
            A.append(distance_PR[i]/(new_P[i] * new_f[i] * new_w_plus_1[i]))
    
    n_360 = [(math.floor(x / 360) * 360) for x in g]
    g_360 = [x - (math.floor(x / 360) * 360) for x in g]

    A.append(A[2]*0.27)
    A.append(A[4]*0.33)
    g_360.append(g_360[2])
    g_360.append(g_360[4])

    # print('distancePR ', [round(x, 3) for x in distance_PR])
    # print('v ', new_v)
    # print('P ', new_P)
    # print('f ', new_f)
    # print('u ', new_u )
    # print('p', new_p)
    # print('rums', rums)
    # print('new_r', new_r)
    # print('new_w', new_w)
    # print('new_w_plus_1', new_w_plus_1)
    # print('g', g)
    # print('n_360', n_360)
    print('A', [round(x,2) for x in A])
    print('g_360', [round(x, 2) for x in g_360])

    formzahl = (A[4]+A[5])/(A[1]+A[2])
    result = ''

    if 0 < formzahl <= 0.25:
        result = 'Harian Ganda'
    elif 0.25 < formzahl <= 1.5:
        result = 'Campuran Condong Harian Ganda'
    elif 1.5 < formzahl <= 3.0:
        result = 'Campuran Condong Harian tunggal'
    else:
        result = 'Harian Tunggal'

    highest = max(max(row) for row in section1)
    lowest = min(min(row) for row in section1)
    difference = abs(highest - lowest)

    return formzahl, result, highest, lowest, difference, section1, A, g_360
