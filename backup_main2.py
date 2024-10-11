import time
import math
import pandas
import ctypes
import locale
import tkinter
from tkinter import ttk
from datetime import datetime
from tkinter import filedialog
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from tkinter import filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


ctypes.windll.shcore.SetProcessDpiAwareness(1)
locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')

app = tkinter.Tk()
app.title('Measurement App')

w = 1600
h = 850
x = int((app.winfo_screenwidth() / 2) - (w / 2))
y = int((app.winfo_screenheight() / 2) - (h / 2))
app.geometry(f'{w}x{h}+{x}+{y}')

graph1_x = 400
graph1_y = 10
graph1_w = 570
graph1_h = 570

graph2_x = 1000
graph2_y = 10
graph2_w = 570
graph2_h = 570

def update_time():
    current_time = time.strftime("%H:%M:%S")
    label_time.config(text=current_time)
    label_time.after(1000, update_time)

def load_logo():
    image = Image.open('logo.png')
    image = image.resize((90, 60))
    logo = ImageTk.PhotoImage(image)
    
    label_logo = tkinter.Label(top_frame, image=logo, bg='#292F36')
    label_logo.image = logo
    label_logo.place(x=10, y=10)

def on_closing():
    if messagebox.askokcancel('Quit', 'Apakah Anda ingin menutup aplikasi?'):
        app.quit()
        app.destroy()

def load_data():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
    entry_path.delete(0, tkinter.END)
    entry_path.insert(0, file_path) 

def create_empty_plot(frame, x, y, width, height):
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.plot([], [])
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.set_xlabel("X-axis")
    ax.set_ylabel("Y-axis")

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().place(x=x, y=y, width=width, height=height)

def create_table(frame, x, y, width, height):
    columns = ('col1', 'col2', 'col3', 'col4', 'col5')
    table = ttk.Treeview(frame, columns=columns, show='headings')

    table.heading('col1', text='No')
    table.heading('col2', text='Tanggal')
    table.heading('col3', text='Jam')
    table.heading('col4', text='Elevasi')
    table.heading('col5', text='Elevasi Filter')

    data = [
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
        ['', '', '','',''],
    ]
    for row in data:
        table.insert('', tkinter.END, values=row)

    scrollbar = ttk.Scrollbar(frame, orient=tkinter.VERTICAL, command=table.yview)
    table.configure(yscroll=scrollbar.set)

    table.place(x=x, y=y, width=width - 20, height=height)
    scrollbar.place(x=x+width-20, y=y, height=height)

    return table

def process_data(file_path):
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

    # for spr in sum_PR:
    #     print([round(sp, 3) for sp in spr])

    distance_PR = []

    for i in range(len(sum_PR[0])):
        x1_squared = sum_PR[1][i] * sum_PR[1][i]
        x0_squared = sum_PR[0][i] * sum_PR[0][i]
        distance = math.sqrt(abs(x1_squared + x0_squared))
        distance_PR.append(distance)

    # print([round(d, 3) for d in distance_PR])

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

    # print([round(di,3) for di in r_distance])

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

    print('w+1', [round(ww, 3) for ww in w_plus_1])

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

    show_data(formzahl, result, highest, lowest, difference, section1)


def show_data(formzahl, result, highest, lowest, difference, section_1, row_result_straight=None, result_table=None):
    entry_height_max.delete(0, tkinter.END)
    entry_height_max.insert(0, highest) 

    entry_height_min.delete(0, tkinter.END)
    entry_height_min.insert(0, lowest) 

    entry_height_diff.delete(0, tkinter.END)
    entry_height_diff.insert(0, difference) 

    entry_formzahl.delete(0, tkinter.END)
    entry_formzahl.insert(0, formzahl) 

    entry_type.delete(0, tkinter.END)
    entry_type.insert(0, result) 
    

    for row in table.get_children():
        table.delete(row)

    fig, ax = plt.subplots(figsize=(4, 3))
    for i, row in enumerate(section_1):
        ax.plot(row, label=f'Line {i+1}')

    ax.set_xlabel('Index')
    ax.set_ylabel('Value')
    ax.set_title('2D Plot')
    ax.legend(loc='best', fontsize='small')

    canvas = FigureCanvasTkAgg(fig, master=mid_frame)
    canvas.draw()
    canvas.get_tk_widget().place(x=graph1_x, y=graph1_y, width=graph1_w, height=graph1_h) 


file_path = None

top_frame_height = int(0.1 * h)
top_frame = tkinter.Frame(app, background='#292F36', height=top_frame_height, width=w)
top_frame.place(x=0, y=0)

mid_frame_height = int(0.83 * h)
mid_frame = tkinter.Frame(app, background='#FAF5F1', height=mid_frame_height, width=w)
mid_frame.place(x=0, y=top_frame_height)

bot_frame_height = int(0.07 * h)
bot_frame = tkinter.Frame(app, background='#E0D8D8', height=bot_frame_height, width=w)
bot_frame.place(x=0, y=top_frame_height+mid_frame_height)

result_frame = tkinter.LabelFrame(mid_frame, text='Hasil', height=300, width=380)
result_frame.place(x=10, y=370)

for i, text in enumerate(['Tinggi Air Max', 'Tinggi Air Min', 'Tunggang Pasut', 'Formzahl', 'Tipe Pasut']):
    tkinter.Label(result_frame, text=text).place(x=10, y=10+40*i)

entry_height_max = tkinter.Entry(result_frame, width=25)
entry_height_max.place(x=140, y=10)
entry_height_min = tkinter.Entry(result_frame, width=25)
entry_height_min.place(x=140, y=50)
entry_height_diff = tkinter.Entry(result_frame, width=25)
entry_height_diff.place(x=140, y=90)
entry_formzahl = tkinter.Entry(result_frame, width=25)
entry_formzahl.place(x=140, y=130)
entry_type = tkinter.Entry(result_frame, width=25)
entry_type.place(x=140, y=170)

entry_path = tkinter.Entry(mid_frame, width=45)
entry_path.place(x=10, y=290)

label_time = tkinter.Label(top_frame, text="", fg='white', bg='#292F36', font=('Helvetica', 16))
label_time.place(x=w-140, y=30)

button_load = tkinter.Button(mid_frame, text='Load Excel', command=load_data)
button_load.place(x=10, y=320)

button_execute = tkinter.Button(mid_frame, text='Jalankan', command=lambda: process_data(file_path))
button_execute.place(x=100, y=320)

button_execute = tkinter.Button(mid_frame, text='Simpan Hasil')
button_execute.place(x=170, y=320)

table = create_table(mid_frame, 10, 10, 380, 230)
create_empty_plot(mid_frame, graph1_x, graph1_y, graph1_w, graph1_h)
create_empty_plot(mid_frame, graph2_x, graph2_y, graph2_w, graph2_h)

tkinter.Label(bot_frame, text='Copyright @', bg='#E0D8D8').place(x=w-110, y=20)


load_logo()
update_time()
app.protocol('WM_DELETE_WINDOW', on_closing)
app.mainloop()
