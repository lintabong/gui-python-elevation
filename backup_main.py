import time
import ctypes
import locale
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from helper import process_data
from constant import *

ctypes.windll.shcore.SetProcessDpiAwareness(2)
locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')

class MeasurementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Measurement App')

        self.dpi = self.winfo_fpixels('1i')

        self.width = int(1280 * (self.dpi / 96))
        self.height = int(700 * (self.dpi / 96))
        x_pos = int((self.winfo_screenwidth() / 2) - (self.width / 2))
        y_pos = int((self.winfo_screenheight() / 2) - (self.height / 2))
        self.geometry(f'{self.width}x{self.height}+{x_pos}+{y_pos}')
        
        self.file_path = None
        self.label_time = None

        # self.create_frames()
        self.top_frame()
        self.mid_frame()
        self.bot_frame()
        # self.load_logo()
        self.update_time()

        # # Add Widgets
        # self.create_widgets()
        
        self.protocol('WM_DELETE_WINDOW', self.on_closing)

    def top_frame(self):
        top_frame = tk.Frame(self, background='#292F36', height=int(0.1 * self.height), width=self.width)
        top_frame.place(x=0, y=0)

        image = Image.open('logo.png')
        image = image.resize((90, 60))
        logo = ImageTk.PhotoImage(image)
        label_logo = tk.Label(top_frame, image=logo, bg='#292F36')
        label_logo.image = logo
        label_logo.place(x=10, y=10)

        self.label_time = tk.Label(top_frame, text="", fg='white', bg='#292F36', font=('Helvetica', 16))
        self.label_time.place(x=1280-140, y=30)

    def update_time(self):
        current_time = time.strftime('%H:%M:%S')
        self.label_time.config(text=current_time)
        self.label_time.after(1000, self.update_time)

    def mid_frame(self):
        mid_frame = tk.Frame(self, background='#FAF5F1', height=int(0.83 * self.height), width=self.width)
        mid_frame.place(x=0, y=int(0.1 * self.height))

    def bot_frame(self):
        bot_frame = tk.Frame(self, background='#E0D8D8', height=int(0.07 * self.height), width=self.width)
        bot_frame.place(x=0, y=int(0.93 * self.height))

    # def create_widgets(self):
    #     self.label_time = tk.Label(self.top_frame, text="", fg='white', bg='#292F36', font=('Helvetica', 16))
    #     self.label_time.place(x=1280-140, y=30)
        
    #     self.entry_path = tk.Entry(self.mid_frame, width=40)
    #     self.entry_path.place(x=10, y=290)

    #     tk.Button(self.mid_frame, text='Load Excel', command=self.load_data).place(x=10, y=320)
    #     tk.Button(self.mid_frame, text='Jalankan', command=self.execute).place(x=100, y=320)
    #     tk.Button(self.mid_frame, text='Simpan Hasil').place(x=170, y=320)

    #     self.create_result_frame()
    #     self.create_table(self.mid_frame, 10, 10, 380, 230)
    #     self.create_constant_table(self.mid_frame, 400, 590, 1050, 100)
    #     self.create_empty_plot(self.mid_frame, 400, 10, 400, 300)
    #     self.create_empty_plot(self.mid_frame, 850, 10, 400, 300)

    # def create_result_frame(self):
    #     self.result_frame = tk.LabelFrame(self.mid_frame, text='Hasil', height=300, width=380)
    #     self.result_frame.place(x=10, y=370)

    #     entries_label = ['Tinggi Air Max', 'Tinggi Air Min', 'Tunggang Pasut', 'Formzahl', 'Tipe Pasut']
    #     self.entries = {}

    #     for i, text in enumerate(entries_label):
    #         tk.Label(self.result_frame, text=text).place(x=10, y=10+40*i)
    #         entry = tk.Entry(self.result_frame, width=20)
    #         entry.place(x=140, y=10+40*i)
    #         self.entries[text] = entry

    # def create_empty_plot(self, frame, x, y, width, height):
    #     fig, ax = plt.subplots(figsize=(3, 2))
    #     ax.plot([], [])
    #     ax.set_xlim(0, 10)
    #     ax.set_ylim(0, 10)
    #     canvas = FigureCanvasTkAgg(fig, master=frame)
    #     canvas.draw()
    #     canvas.get_tk_widget().place(x=x, y=y, width=width, height=height)

    # def create_table(self, frame, x, y, width, height):
    #     columns = ('col1', 'col2', 'col3', 'col4', 'col5')
    #     table = ttk.Treeview(frame, columns=columns, show='headings')

    #     for col in columns:
    #         table.heading(col, text=col)

    #     table.place(x=x, y=y, width=width-20, height=height)
    #     scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=table.yview)
    #     table.configure(yscroll=scrollbar.set)
    #     scrollbar.place(x=x+width-20, y=y, height=height)
    #     return table

    # def create_constant_table(self, frame, x, y, width, height):
    #     columns = ('Konstanta', 'S0', 'M2', 'S2', 'N2', 'K1', 'Q1', 'M4', 'MS4', 'K2', 'P1')
    #     table = ttk.Treeview(frame, columns=columns, show='headings')

    #     for col in columns:
    #         table.heading(col, text=col)
    #         table.column(col, width=100)

    #     table.place(x=x, y=y, width=width-20, height=height)
    #     scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=table.yview)
    #     table.configure(yscroll=scrollbar.set)
    #     scrollbar.place(x=x+width-20, y=y, height=height)
    #     return table

    def load_logo(self):
        image = Image.open('logo.png')
        image = image.resize((90, 60))
        logo = ImageTk.PhotoImage(image)
        label_logo = tk.Label(self.top_frame, image=logo, bg='#292F36')
        label_logo.image = logo
        label_logo.place(x=10, y=10)

    def update_time(self):
        current_time = time.strftime('%H:%M:%S')
        self.label_time.config(text=current_time)
        self.label_time.after(1000, self.update_time)

    # def load_data(self):
    #     self.file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
    #     self.entry_path.delete(0, tk.END)
    #     self.entry_path.insert(0, self.file_path)

    # def execute(self):
    #     if self.file_path:
    #         formzahl, result, highest, lowest, difference, section1, H, g_360 = process_data.run(file_path=self.file_path)
    #         new_data = [['A cm'], ['g360']]

    #         for i in range(len(H)):
    #             new_data[0].append(round(H[i], 3))
    #             new_data[1].append(round(g_360[i], 3))

    #         new_data[0].append(round(H[2] * 0.27, 3))
    #         new_data[0].append(round(H[4] * 0.33, 3))

    #         self.entries['Tinggi Air Max'].delete(0, tk.END)
    #         self.entries['Tinggi Air Max'].insert(0, highest)
    #         self.entries['Tinggi Air Min'].delete(0, tk.END)
    #         self.entries['Tinggi Air Min'].insert(0, lowest)
    #         self.entries['Tunggang Pasut'].delete(0, tk.END)
    #         self.entries['Tunggang Pasut'].insert(0, difference)
    #         self.entries['Formzahl'].delete(0, tk.END)
    #         self.entries['Formzahl'].insert(0, formzahl)
    #         self.entries['Tipe Pasut'].delete(0, tk.END)
    #         self.entries['Tipe Pasut'].insert(0, result)

    def on_closing(self):
        self.quit()
        self.destroy()

if __name__ == "__main__":
    app = MeasurementApp()
    app.mainloop()
