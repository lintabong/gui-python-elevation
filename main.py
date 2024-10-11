import time
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

from helper import process_data


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
    current_time = time.strftime('%H:%M:%S')
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

def execute(file_path):
    formzahl, result, highest, lowest, difference, section1, H, g_360 = process_data.run(file_path=file_path)

    new_data = [['A cm'],['g360']]
    
    for i in range(len(H)):
        new_data[0].append(round(H[i], 3))
        new_data[1].append(round(g_360[i], 3))

    new_data[0].append(round(H[2]*0.27, 3))
    new_data[0].append(round(H[4]*0.33, 3))
    new_data[1].append(round(g_360[2], 3))
    new_data[1].append(round(g_360[4], 3))

    update_constant_table(constant_table, new_data)
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
    for i, row in enumerate(section1):
        ax.plot(row, label=f'Line {i+1}')

    ax.set_xlabel('Index')
    ax.set_ylabel('Value')
    ax.set_title('2D Plot')
    ax.legend(loc='best', fontsize='small')

    canvas = FigureCanvasTkAgg(fig, master=mid_frame)
    canvas.draw()
    canvas.get_tk_widget().place(x=graph1_x, y=graph1_y, width=graph1_w, height=graph1_h) 


def create_constant_table(frame, x, y, width, height):
    columns = ('Konstanta', 'S0', 'M2', 'S2', 'N2', 'K1', 'Q1', 'M4', 'MS4', 'K2', 'P1')
    table = ttk.Treeview(frame, columns=columns, show='headings')

    table.heading('Konstanta', text='Konstanta')
    table.heading('S0', text='S0')
    table.heading('M2', text='M2')
    table.heading('S2', text='S2')
    table.heading('N2', text='N2')
    table.heading('K1', text='K1')
    table.heading('Q1', text='Q1')
    table.heading('M4', text='M4')
    table.heading('MS4', text='MS4')
    table.heading('K2', text='K2')
    table.heading('P1', text='P1')

    column_width = 80
    table.column('Konstanta', width=100)
    table.column('S0', width=column_width)
    table.column('M2', width=column_width)
    table.column('S2', width=column_width)
    table.column('N2', width=column_width)
    table.column('K1', width=column_width)
    table.column('Q1', width=column_width)
    table.column('M4', width=column_width)
    table.column('MS4', width=column_width)
    table.column('K2', width=column_width)
    table.column('P1', width=column_width)

    # Initial data for the table
    data = [
        ['A cm', '', '', '', '', '', '', '', '', '', ''],
        ['g360', '', '', '', '', '', '', '', '', '', '']
    ]
    
    for row in data:
        table.insert('', tkinter.END, values=row)

    scrollbar = ttk.Scrollbar(frame, orient=tkinter.VERTICAL, command=table.yview)
    table.configure(yscroll=scrollbar.set)

    table.place(x=x, y=y, width=width - 20, height=height)
    scrollbar.place(x=x+width-20, y=y, height=height)

    return table

def update_constant_table(table, new_data):
    for row in table.get_children():
        table.delete(row)

    for row in new_data:
        table.insert('', tkinter.END, values=row)

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

button_execute = tkinter.Button(mid_frame, text='Jalankan', command=lambda: execute(file_path))
button_execute.place(x=100, y=320)

button_execute = tkinter.Button(mid_frame, text='Simpan Hasil')
button_execute.place(x=170, y=320)

table = create_table(mid_frame, 10, 10, 380, 230)

constant_table = create_constant_table(mid_frame, 400, 590, 1050, 100)
create_empty_plot(mid_frame, graph1_x, graph1_y, graph1_w, graph1_h)
create_empty_plot(mid_frame, graph2_x, graph2_y, graph2_w, graph2_h)

tkinter.Label(bot_frame, text='Copyright @', bg='#E0D8D8').place(x=w-110, y=20)


load_logo()
update_time()
app.protocol('WM_DELETE_WINDOW', on_closing)
app.mainloop()
