import math
import pandas
from datetime import datetime


def run(file_path):
    df = pandas.read_excel(file_path)
    df_clean = (df.iloc[10:41, 1:26]).dropna(how='all')

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

    # section3
    cx_1 = 10
    cy_1 = 10
    cx_2 = 10
    cy_2 = 10
    cx_4 = 10
    cy_4 = 10

    section3 = [
        [
            sum(row[0:2]),
            row[0] - row[1] + cx_1,
            row[2] - row[3] + cy_1,
            row[4] - row[5] + cx_2,
            row[6] - row[7] + cy_2,
            row[8] - row[9] + cx_4,
            row[10] - row[11] + cy_4
        ]
        for row in section2
    ]

    # section 4
    suffix_0_X = 0
    suffix_10_X = 0
    suffix_10_Y = 0
    suffix_12_X = 0
    suffix_12_Y = 0
    suffix_1b_X = 0
    suffix_1b_Y = 0
    suffix_13_X = 0
    suffix_13_Y = 0

    suffix_1c_X = 0
    suffix_1c_Y = 0
    suffix_20_X = 0
    suffix_20_Y = 0

    suffix_22_X = 0
    suffix_22_Y = 0

    suffix_2b_X = 0
    suffix_2b_Y = 0
    suffix_23_X = 0
    suffix_23_Y = 0

    suffix_2c_X = 0
    suffix_2c_Y = 0
    suffix_42_X = 0
    suffix_42_Y = 0
    suffix_4b_X = 0
    suffix_4b_Y = 0
    suffix_44_X = 0
    suffix_44_Y = 0

    suffix_4d_X = 0
    suffix_4d_Y = 0

    for i, row in enumerate(section3):
        suffix_0_X += row[0]
        suffix_10_X += row[1]
        suffix_10_Y += row[2]
        suffix_20_X += row[3]
        suffix_20_Y += row[4]

        if i in [0,1,2,3,11,12,13,14,15,16,17,25,26,27,28]:
            suffix_12_X += row[1]
            suffix_12_Y += row[2]
            suffix_22_X += row[3]
            suffix_22_Y += row[4]
            suffix_42_X += row[5]
            suffix_42_Y += row[6]
        else:
            suffix_12_X -= row[1]
            suffix_12_Y -= row[2]
            suffix_22_X -= row[3]
            suffix_22_Y -= row[4]
            suffix_42_X -= row[5]
            suffix_42_Y -= row[6]

        if i in [8,9,10,11,12,13,22,23,24,25,26,27]:
            suffix_1b_X += row[1]
            suffix_1b_Y += row[2]
            suffix_2b_X += row[3]
            suffix_2b_Y += row[4]
            suffix_4b_X += row[5]
            suffix_4b_Y += row[6]
        
        if i in [1,2,3,4,5,6,15,16,17,18,19,20]:
            suffix_1b_X -= row[1]
            suffix_1b_Y -= row[2]
            suffix_2b_X -= row[3]
            suffix_2b_Y -= row[4]
            suffix_4b_X -= row[5]
            suffix_4b_Y -= row[6]
        
        if i in [2,3,4,5,6,12,13,14,15,16,22,23,24,25,26]:
            suffix_13_X += row[1]
            suffix_13_Y += row[2]
            suffix_23_X += row[3]
            suffix_23_Y += row[4]
        else:
            suffix_13_X -= row[1]
            suffix_13_Y -= row[2]
            suffix_23_X -= row[3]
            suffix_23_Y -= row[4]

        if i in [0,1,2,3,4,10,11,12,13,19,20,21,22,23]:
            suffix_1c_X += row[1]
            suffix_1c_Y += row[2]
            suffix_2c_X += row[3]
            suffix_2c_Y += row[4]
        
        if i in [5,6,7,8,9,15,16,17,18,24,25,26,27,28]:
            suffix_1c_X -= row[1]
            suffix_1c_Y -= row[2]
            suffix_2c_X -= row[3]
            suffix_2c_Y -= row[4]

        if i in [0,1,5,6,7,8,13,14,15,20,21,22,23,27,28]:
            suffix_44_X += row[5]
            suffix_44_Y += row[6]
        if i in [2,3,4,9,10,11,12,16,17,18,19,24,25,26]:
            suffix_44_X -= row[5]
            suffix_44_Y -= row[6]

        if i in [4,5,6,11,12,13,18,19,20,25,26,27]:
            suffix_4d_X += row[5]
            suffix_4d_Y += row[6]
        if i in [1,2,3,8,9,10,15,16,17,22,23,24]:
            suffix_4d_X -= row[5]
            suffix_4d_Y -= row[6]

    suffix_10_X -= 29*cx_1
    suffix_10_Y -= 29*cy_1
    suffix_20_X -= 29*cx_2
    suffix_20_Y -= 29*cy_2

    suffix_12_X -= 1*cx_1
    suffix_12_Y -= 1*cy_1
    suffix_22_X -= 1*cx_2
    suffix_22_Y -= 1*cy_2
    suffix_42_X -= 1*cx_4
    suffix_42_Y -= 1*cy_4

    suffix_13_X -= 1*cx_1
    suffix_13_Y -= 1*cy_1
    suffix_23_X -= 1*cx_2
    suffix_23_Y -= 1*cy_2

    suffix_44_X -= 1*cx_4
    suffix_44_Y -= 1*cy_4

    # section 5
    constant5 = [suffix_0_X, suffix_10_X, suffix_12_X - suffix_1b_Y, suffix_13_X - suffix_1c_Y, suffix_20_X, suffix_22_X - suffix_2b_Y,
                suffix_23_X - suffix_2c_Y, suffix_42_X - suffix_4b_Y, suffix_44_X - suffix_4d_Y]
    section5 = [
        [constant5[0], 0, 0, 0, 0, 0, 0, 0],
        [0, 0, 0, 0, 1*constant5[1], -0.08*constant5[1], 0, 0],
        [0, 0.07*constant5[2], 0, 0, -0.02*constant5[2], 1*constant5[2], 0, 0.02*constant5[2]],
        [0, 0, 0, 0, 0, 0, 0, 0],
        [0, -0.03*constant5[4], 1*constant5[4], -0.03*constant5[4], 0, 0, 0, 0],
        [0, 1*constant5[5], 0.015*constant5[5], 0.038*constant5[5], 0.002*constant5[5], -0.058*constant5[5], 0, -0.035*constant5[5]],
        [0, -0.06*constant5[6], 0, 1*constant5[6], 0, 0, 0, 0],
        [0, 0.03*constant5[7], 0, 0, 0, 0, 0, 1*constant5[7]],
        [0, 0, 0, 0, 0, 0, 1*constant5[8], 0.08*constant5[8]]
    ]

    # section 6
    constant6 = [suffix_10_Y, suffix_12_Y + suffix_1b_X, suffix_13_Y + suffix_1c_X, suffix_20_Y,
                suffix_22_Y + suffix_2b_X, suffix_23_Y + suffix_2c_X, suffix_42_Y + suffix_4b_X,
                suffix_44_Y + suffix_4d_X]

    section6 = [
        [0, 0, 0, 0, 1.01*constant6[0], -0.08*constant6[0], 0, 0],
        [0, 0.07*constant6[1], 0, 0, -0.02*constant6[1], 1*constant6[1], 0, 0.03*constant6[1]],
        [0, 0, 0, 0, 0, 0, 0, 0],
        [0, -0.03*constant6[3], 1*constant6[3], -0.03*constant6[3], 0, 0, 0, 0],
        [0, 1*constant6[4], 0.015*constant6[4], 0.032*constant6[4], 0, -0.057*constant6[4], 0, -0.035*constant6[4]],
        [0, -0.06*constant6[5], 0, 1*constant6[5], 0, 0, 0, 0],
        [0, 0.03*constant6[6], 0, 0, 0, 0, 0.01*constant6[6], 1*constant6[6]],
        [0, 0, 0, 0, 0, 0, 1*constant6[7], 0.08*constant6[7]]
    ]

    # for row in section6:
    #     print([round(r, 3) for r in row])
    # section 7

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

    f_M2 = 0.976
    u_M2 = 1.6
    P = [696,	559,	448,	566,	439,	565,	507,	535]
    f = [0, f_M2,	1.000,	f_M2,	1.082,	1.132,	f_M2*f_M2,	f_M2]

    E0_1 = [0, 250.1, 0, 4, 10.8, 239.3, 0, 0]
    E0_2 = [0, 231.1, 0, 341.3, 209, 22.2, 0, 0]
    E0_3 = [0, 18.700, 0, 195.700, 13.800, 4.900, 0, 0]

    E = [E0_1[i] + E0_2[i] + E0_3[i] for i in range(len(E0_1))]
    u = [0, u_M2, 0, u_M2, 6.100, -6.900, u_M2*2, u_M2]

    p = [0, 333, 345, 327, 173, 160, 307, 318]

    r_distance = [distance_PR[0]]
    for i in range(len(distance_PR)):
        if i > 0:
            r_distance.append((distance_PR[i] / sum_PR[1][i] ))

    r = [
        0,
        ((r_distance[1] - (-4.01)) / ((-4.33) - (-4.01))) * (103 - 104) + 104,
        ((r_distance[2] - 0.7) / (0.727 - 0.7)) * (36 - 35) + 35,
        ((r_distance[3] - (-0.052)) / (-0.07 - (-0.052))) * (176 - 177) + 176,
        ((r_distance[4] - (-1)) / (-1.036 - (-1))) * (134 - 135) + 135,
        ((r_distance[5] - (-0.268)) / (-0.287 - (-0.268))) * (344 - 345) + 345,
        ((r_distance[6] - (-2.14)) / (-2.25 - (-2.14))) * (294 - 295) + 295,
        ((r_distance[7] - (-1.192)) / (-1.235 - (-1.192))) * (309 - 310) + 310
    ]

    # section 9
    # w and 1+W for S2, MS4
    E_s2 = E[4]
    u_s2 = u[4]
    Eu_s2 = E_s2 + u_s2
    wf_s2 = ((Eu_s2 - 230) / (240 - 230)) * (-9.8 - (-13)) + (-13)
    Wf_s2 = ((Eu_s2 - 230) / (240 - 230)) * (0.115 - 0.029) + 0.029
    f_s2 = 1.212
    w_s2 = wf_s2 * f_s2
    W_s2 = Wf_s2 * f_s2
    W_plus_1_s2 = 1 + W_s2

    Edot2_k1 = 2*E[4]
    u_k1 = u[4]
    Eu_k1 = (Edot2_k1 + u_k1)-360
    wf_k1 = ((Eu_k1 - 110) / (120 - 110)) * (-19 - (-19.3)) + (-19.3)
    Wf_k1 = ((Eu_k1 - 110) / (120 - 110)) * (-0.118 - (-0.06)) + (-0.06)
    f_k1 = f[4]
    w_k1 = wf_k1/f_k1
    W_k1 = Wf_k1/f_k1
    W_plus_1_k1 = 1 + W_k1

    Edot3_n2 = 3*E[1]
    Edot2_n2 = 2*E[3]
    diff_E_n2 = abs(Edot3_n2 - Edot2_n2)-360
    w_n2 = ((diff_E_n2 - 80) / (60 - 50)) * (8.3 - 7.2) + 7.2
    w_plus_1_n2 = ((diff_E_n2 - 80) / (90 - 80)) * (1.017 - 1.048) + 1.048

    w = [0, 0, w_s2, w_n2, w_k1, 0, 0, w_s2]
    w_plus_1 = [0, w_plus_1_n2, W_plus_1_s2, w_plus_1_n2, W_plus_1_k1, 1, 1, W_plus_1_s2]

    g = [E[i] + u[i] + w[i] + p[i] + r[i] for i in range(len(w))]

    g_360 = [g[i] % 360 for i in range(len(g))]
    H = []

    for i in range(len(distance_PR)):
        if i == 0:
            H.append(distance_PR[i]/P[i])
        else:
            H.append(distance_PR[i]/(P[i]*w_plus_1[i]*f[i]))

    formzahl = (H[4]+H[5])/(H[1]+H[2])
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

    return formzahl, result, highest, lowest, difference, section1, H, g_360
