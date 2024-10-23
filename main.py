import time
import ctypes
import locale
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from helper import process_data, export_excel
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

        self.top_frame()
        self.mid_frame()
        self.bot_frame()
        
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

        label_time = tk.Label(top_frame, text="", fg='white', bg='#292F36', font=('Helvetica', 16))
        label_time.place(x=self.width-140, y=30)

        def update_time():
            current_time = time.strftime('%H:%M:%S')
            label_time.config(text=current_time)
            label_time.after(1000, update_time)

        update_time()

    def mid_frame(self):
        global file_path, formzahl, result, highest, lowest, difference, section1, H, g_360

        entries_label = ['Tinggi Air Max', 'Tinggi Air Min', 'Tunggang Pasut', 'Formzahl', 'Tipe Pasut']
        entries = {}

        def load_data():
            global file_path
            file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
            entry_path.delete(0, tk.END)
            entry_path.insert(0, file_path)

        def execute():
            global file_path, formzahl, result, highest, lowest, difference, section1, H, g_360

            if file_path:
                formzahl, result, highest, lowest, difference, section1, H, g_360 = process_data.run(file_path=file_path)

                result_process = [highest, lowest, difference, formzahl, result]

                for i, text in enumerate(entries_label):
                    entries[text].delete(0, tk.END)
                    entries[text].insert(0, result_process[i])

                create_plot(mid_frame, 430, 10, 550, 550, section1, f'Plot {str(len(section1))} hari')
                create_plot(mid_frame, 1010, 10, 550, 550, [section1[len(section1)//2]], 'Plot tengah')

                new_data = [['A cm'],['g360']]
    
                for i in range(len(H)):
                    new_data[0].append(round(H[i], 3))
                    new_data[1].append(round(g_360[i], 3))

                for row in table.get_children():
                    table.delete(row)

                for row in new_data:
                    table.insert('', tk.END, values=row)

                if section1:
                    dropdown_var.set('')
                    dropdown_picker['values'] = list(range(1, len(section1) + 1))

        def save_result():
            global file_path, formzahl, result, highest, lowest, difference, section1, H, g_360

            wb = export_excel.run(formzahl, result, highest, lowest, difference, section1, H, g_360)
            
            save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
            if save_path:
                wb.save(save_path)
                messagebox.showinfo('Success', f'Data berhasil disimpan ke {save_path}')
            else:
                messagebox.showwarning('Cancelled', 'Penyimpanan dibatalkan')

        def create_plot(frame, x, y, width, height, data_y = [[]], title=None):           
            fig, ax = plt.subplots(figsize=(4, 3))
            for i, row in enumerate(data_y):
                ax.plot(row)

            if len(data_y) <= 0:
                ax.set_xlim(0, 10)
                ax.set_ylim(0, 10)

            ax.set_xlabel('Index')
            ax.set_ylabel('Value')
            ax.set_title('2D Plot') if not title else ax.set_title(title)
            ax.legend(loc='best', fontsize='small')

            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().place(x=x, y=y, width=width, height=height)

        mid_frame = tk.Frame(self, background='#FAF5F1', height=int(0.83 * self.height), width=self.width)
        mid_frame.place(x=0, y=int(0.1 * self.height))

        entry_path = tk.Entry(mid_frame, width=45)
        entry_path.place(x=10, y=10)

        tk.Button(mid_frame, text='Load Excel', command=load_data).place(x=10, y=60)
        tk.Button(mid_frame, text='Jalankan', command=execute).place(x=100, y=60)
        tk.Button(mid_frame, text='Simpan Hasil', command=save_result).place(x=170, y=60)

        columns = ('Konstanta', 'S0', 'M2', 'S2', 'N2', 'K1', 'Q1', 'M4', 'MS4', 'K2', 'P1')
        table = ttk.Treeview(mid_frame, columns=columns, show='headings')

        result_frame = tk.LabelFrame(mid_frame, text='Hasil', height=300, width=380)
        result_frame.place(x=10, y=240)

        for i, text in enumerate(entries_label):
            tk.Label(result_frame, text=text).place(x=10, y=10+40*i)
            entry = tk.Entry(result_frame, width=20)
            entry.place(x=140, y=10+40*i)
            entries[text] = entry

        x = 440
        y = 580
        table_width = 1100
        for col in columns:
            table.heading(col, text=col)
            table.column(col, width=10)

        table.place(x=x, y=y, width=table_width-20, height=130)
        scrollbar = ttk.Scrollbar(mid_frame, orient=tk.VERTICAL, command=table.yview)
        table.configure(yscroll=scrollbar.set)
        scrollbar.place(x=x+table_width-20, y=y, height=230)
        
        controll_frame = tk.LabelFrame(mid_frame, text='Plot Hari tertentu', height=100, width=380)
        controll_frame.place(x=10, y=120)

        dropdown_var = StringVar()
        dropdown_picker = ttk.Combobox(controll_frame, textvariable=dropdown_var, state="readonly")
        dropdown_picker.place(x=40, y=10)

        def on_dropdown_change(event):
            selected_day = int(dropdown_var.get()) - 1
            create_plot(mid_frame, 1010, 10, 550, 550, [section1[selected_day]], f'Plot hari ke-{selected_day + 1}')

        dropdown_picker.bind("<<ComboboxSelected>>", on_dropdown_change)

        create_plot(mid_frame, 430, 10, 550, 550)
        create_plot(mid_frame, 1010, 10, 550, 550)

    def bot_frame(self):
        bot_frame = tk.Frame(self, background='#E0D8D8', height=int(0.07 * self.height), width=self.width)
        bot_frame.place(x=0, y=int(0.93 * self.height))

    def on_closing(self):
        self.quit()
        self.destroy()

if __name__ == "__main__":
    app = MeasurementApp()
    app.mainloop()
