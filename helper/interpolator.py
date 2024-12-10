def interpolate1_0(z8, const_w_1):
    col = 1
    if 0 <= z8 <= 10:
        index = 0
    elif 10 < z8 <= 20:
        index = 1
    elif 20 < z8 <= 30:
        index = 2
    elif 30 < z8 <= 40:
        index = 3
    elif 40 < z8 <= 50:
        index = 4
    elif 50 < z8 <= 60:
        index = 5
    else:
        return None
    return const_w_1[index][col] + ((z8 - 0) / 10) * (const_w_1[index+1][col] - const_w_1[index][col])


def interpolate1_1(z8, const_w_1):
    col = 1
    if 60 < z8 <= 70:
        index = 6
    elif 70 < z8 <= 80:
        index = 7
    elif 80 < z8 <= 90:
        index = 8
    elif 90 < z8 <= 100:
        index = 9
    elif 100 < z8 <= 110:
        index = 10
    elif 110 < z8 <= 120:
        index = 11
    else:
        return None
    return const_w_1[index][col] + ((z8 - 0) / 10) * (const_w_1[index+1][col] - const_w_1[index][col])

def interpolate1_2(z8, const_w_1):
    col = 1
    if 120 < z8 <= 130:
        index = 12
    elif 130 < z8 <= 140:
        index = 13
    elif 140 < z8 <= 150:
        index = 14
    elif 150 < z8 <= 160:
        index = 15
    elif 160 < z8 <= 170:
        index = 16
    elif 170 < z8 <= 180:
        index = 17
    else:
        return None
    return const_w_1[index][col] + ((z8 - 0) / 10) * (const_w_1[index+1][col] - const_w_1[index][col])

def interpolate1_3(z8, const_w_1):
    col = 1
    if 180 < z8 <= 190:
        index = 18
    elif 190 < z8 <= 200:
        index = 19
    elif 200 < z8 <= 210:
        index = 20
    elif 210 < z8 <= 220:
        index = 21
    elif 220 < z8 <= 230:
        index = 22
    elif 230 < z8 <= 240:
        index = 23
    else:
        return None
    return const_w_1[index][col] + ((z8 - 0) / 10) * (const_w_1[index+1][col] - const_w_1[index][col])

def interpolate1_4(z8, const_w_1):
    col = 1
    for low in range(240, 301, 10):
        if low < z8 <= low + 10:
            index = (low - 240) // 10 + 24
            return const_w_1[index][col] + ((z8 - low) / 10) * (const_w_1[index + 1][col] - const_w_1[index][col])
    return None


def interpolate_0(z8, const_w_1, col, low_threshold, high_threshold):
    for low in range(low_threshold, high_threshold+1, 10):
        if low < z8 <= low + 10:
            index = (low - low_threshold) // 10 + 24
            return const_w_1[index][col] + ((z8 - low) / 10) * (const_w_1[index + 1][col] - const_w_1[index][col])
    return None

def interpolate_ac(z18, const_w_1):
    if 120 < z18 <= 130:
        return const_w_1[0][4] + ((z18 - 120) / 10) * (const_w_1[1][4] - const_w_1[0][4])
    elif 130 < z18 <= 140:
        return const_w_1[1][4] + ((z18 - 130) / 10) * (const_w_1[2][4] - const_w_1[1][4])
    elif 140 < z18 <= 150:
        return const_w_1[2][4] + ((z18 - 140) / 10) * (const_w_1[3][4] - const_w_1[2][4])
    elif 150 < z18 <= 160:
        return const_w_1[3][4] + ((z18 - 150) / 10) * (const_w_1[4][4] - const_w_1[3][4])
    elif 160 < z18 <= 170:
        return const_w_1[4][4] + ((z18 - 160) / 10) * (const_w_1[5][4] - const_w_1[4][4])
    elif 170 < z18 <= 180:
        return const_w_1[5][4] + ((z18 - 170) / 10) * (const_w_1[6][4] - const_w_1[5][4])
    else:
        return None
    
def interpolate_ad_range(z18, const_w_1):
    if 120 < z18 <= 130:
        return const_w_1[12][5] + ((z18 - 120) / 10) * (const_w_1[13][5] - const_w_1[12][5])
    elif 130 < z18 <= 140:
        return const_w_1[13][5] + ((z18 - 130) / 10) * (const_w_1[14][5] - const_w_1[13][5])
    elif 140 < z18 <= 150:
        return const_w_1[14][5] + ((z18 - 140) / 10) * (const_w_1[15][5] - const_w_1[14][5])
    elif 150 < z18 <= 160:
        return const_w_1[15][5] + ((z18 - 150) / 10) * (const_w_1[16][5] - const_w_1[15][5])
    elif 160 < z18 <= 170:
        return const_w_1[16][5] + ((z18 - 160) / 10) * (const_w_1[17][5] - const_w_1[16][5])
    elif 170 < z18 <= 180:
        return const_w_1[17][5] + ((z18 - 170) / 10) * (const_w_1[18][5] - const_w_1[17][5])
    else:
        return None

def interpolate_1(z8, const_w_1, col, low_threshold, high_threshold):
    for low in range(low_threshold, high_threshold+1, 10):
        if low < z8 <= low + 10:
            index = (low - low_threshold) // 10 + 12
            return const_w_1[index][col] + ((z8 - low) / 10) * (const_w_1[index + 1][col] - const_w_1[index][col])
    return None

def interpolate_2(z28, const_w_1, col, low_threshold, high_threshold):
    for low in range(low_threshold, high_threshold + 1, 10):
        if low < z28 <= low + 10:
            index = (low - low_threshold) // 10 + 6
            return const_w_1[index][col] + ((z28 - low) / 10) * (const_w_1[index + 1][col] - const_w_1[index][col])
    return None

def interpolate1_5(z8, const_w_1):
    col = 1
    if 300 < z8 <= 310:
        index = 30
    elif 310 < z8 <= 320:
        index = 31
    elif 320 < z8 <= 330:
        index = 32
    elif 330 < z8 <= 340:
        index = 33
    elif 340 < z8 <= 350:
        index = 34
    elif 350 < z8 <= 360:
        index = 35
    else:
        return None
    return const_w_1[index][col] + ((z8 - 0) / 10) * (const_w_1[index+1][col] - const_w_1[index][col])
