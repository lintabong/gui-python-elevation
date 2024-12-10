import os
import json
import time
import ctypes
import locale
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from helper import process_data_2, export_excel

ctypes.windll.shcore.SetProcessDpiAwareness(2)
locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')

class MeasurementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Measurement App')

        self.dpi = self.winfo_fpixels('1i')

        self.width = int(1280 * (self.dpi / 96))
        self.height = int(800 * (self.dpi / 96))
        x_pos = int((self.winfo_screenwidth() / 2) - (self.width / 2))
        y_pos = int((self.winfo_screenheight() / 2) - (self.height / 2))
        self.geometry(f'{self.width}x{self.height}+{x_pos}+{y_pos}')

        self.config_file = 'config.json'
        self.widget_positions = self.load_or_create_config()

        self.top_frame()
        self.mid_frame()
        self.bot_frame()
        
        self.protocol('WM_DELETE_WINDOW', self.on_closing)

    def load_or_create_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                return json.load(f)
        if True:
            widget_positions = {
                'graph_1_x': int(0.27 * self.width),
                'graph_1_y': int(0.015 * self.height),
                'graph_1_width': int(0.72 * self.width),
                'graph_1_height': int(0.32 * self.height),
                'graph_2_x': int(0.27 * self.width),
                'graph_2_y': int(0.35 * self.height),
                'graph_2_width': int(0.72 * self.width),
                'graph_2_height': int(0.32 * self.height),
                'graph_length': int(0.62 * self.height),
                'table_x': int(0.27 * self.width),
                'table_y': int(0.69 * self.height),
                'result_frame_x': 10,
                'result_frame_y': int(0.274 * self.height),
                'controll_frame_x': 10,
                'controll_frame_y': 120,
                'entry_path_length': 20,
                'entry_path_x': 10,
                'entry_path_y': 10,
                'button_y': 60,
                'button_load_x': 10,
                'button_execute_x':100,
                'button_save_x': 170,
                'table_width': 1100,
                'controll_frame_x': 10,
                'controll_frame_y': 120,
                'controll_frame_h': 100,
                'controll_frame_w': 380,
                'result_frame_x': 10,
                'result_frame_y': 240,
                'result_frame_h': 300,
                'result_frame_w': 380,
                'entry_result_length': 24,
                'entry_result_x': 140,
                'entry_result_y': 10,
                'text_result_x': 10,
                'text_result_y': 10
            }
            with open(self.config_file, 'w') as f:
                json.dump(widget_positions, f, indent=4)
            return widget_positions

    def save_widget_positions(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.widget_positions, f, indent=4)

    def top_frame(self):
        top_frame = tk.Frame(self, background='#292F36', height=int(0.1 * self.height), width=self.width)
        top_frame.pack(fill='x', side='top')

        image = Image.open('logo.png')
        image = image.resize((int(0.08 * self.width), int(0.08 * self.height)))
        logo = ImageTk.PhotoImage(image)
        label_logo = tk.Label(top_frame, image=logo, bg='#292F36')
        label_logo.image = logo
        label_logo.pack(side='left', padx=10, pady=10)

        label_time = tk.Label(top_frame, text="", fg='white', bg='#292F36', font=('Helvetica', 16))
        label_time.pack(side='right', padx=10, pady=10)

        def update_time():
            current_time = time.strftime('%H:%M:%S')
            label_time.config(text=current_time)
            label_time.after(1000, update_time)

        update_time()

    def mid_frame(self):
        global file_path, formzahl, result, highest, lowest, difference, section1, H, g_360

        entries_label = ['Tinggi Air Max', 'Tinggi Air Min', 'Tunggang Pasut', 'Formzahl', 'Tipe Pasut']
        entries = {}

        graph_1_x = self.widget_positions['graph_1_x']
        graph_1_y = self.widget_positions['graph_1_y']
        graph_1_w = self.widget_positions['graph_1_width']
        graph_1_h = self.widget_positions['graph_1_height']
        graph_2_x = self.widget_positions['graph_2_x']
        graph_2_y = self.widget_positions['graph_2_y']
        graph_2_w = self.widget_positions['graph_2_width']
        graph_2_h = self.widget_positions['graph_2_height']
        table_x = self.widget_positions['table_x']
        table_y = self.widget_positions['table_y']
        entry_path_length = self.widget_positions['entry_path_length']
        entry_path_x = self.widget_positions['entry_path_x']
        entry_path_y = self.widget_positions['entry_path_y']
        button_y = self.widget_positions['button_y']
        button_load_x = self.widget_positions['button_load_x']
        button_execute_x = self.widget_positions['button_execute_x']
        button_save = self.widget_positions['button_save_x']
        table_width = self.widget_positions['table_width']
        controll_frame_x = self.widget_positions['controll_frame_x']
        controll_frame_y = self.widget_positions['controll_frame_y']
        controll_frame_h = self.widget_positions['controll_frame_h']
        controll_frame_w = self.widget_positions['controll_frame_w']
        result_frame_x = self.widget_positions['result_frame_x']
        result_frame_y = self.widget_positions['result_frame_y']
        result_frame_h = self.widget_positions['result_frame_h']
        result_frame_w = self.widget_positions['result_frame_w']
        entry_result_length = self.widget_positions['entry_result_length']
        entry_result_x = self.widget_positions['entry_result_x']
        entry_result_y = self.widget_positions['entry_result_y']
        text_result_x = self.widget_positions['text_result_x']
        text_result_y = self.widget_positions['text_result_y']

        def load_data():
            global file_path
            file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
            entry_path.delete(0, tk.END)
            entry_path.insert(0, file_path)

        def execute():
            global file_path, formzahl, result, highest, lowest, difference, section1, H, g_360

            if file_path:
                formzahl, result, highest, lowest, difference, section1, H, g_360 = process_data_2.run(file_path=file_path)

                result_process = [highest, lowest, difference, formzahl, result]

                for i, text in enumerate(entries_label):
                    entries[text].delete(0, tk.END)
                    entries[text].insert(0, result_process[i])

                create_plot(mid_frame, graph_1_x, graph_1_y, graph_1_w, graph_1_h, section1, f'Plot {str(len(section1))} hari')
                create_plot(mid_frame, graph_2_x, graph_2_y, graph_2_w, graph_2_h, [section1[len(section1)//2]], 'Plot tengah')

                new_data = [['A cm'],['g°']]
    
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

        entry_path = tk.Entry(mid_frame, width=entry_path_length)
        entry_path.place(x=entry_path_x, y=entry_path_y)

        tk.Button(mid_frame, text='Load Excel', command=load_data).place(x=button_load_x, y=button_y)
        tk.Button(mid_frame, text='Jalankan', command=execute).place(x=button_execute_x, y=button_y)
        tk.Button(mid_frame, text='Simpan Hasil', command=save_result).place(x=button_save, y=button_y)

        columns = ('Konstanta', 'S0', 'M2', 'S2', 'N2', 'K1', 'O1', 'M4', 'MS4', 'K2', 'P1')
        table = ttk.Treeview(mid_frame, columns=columns, show='headings')

        result_frame = tk.LabelFrame(mid_frame, text='Hasil', height=result_frame_h, width=result_frame_w)
        result_frame.place(x=result_frame_x, y=result_frame_y)

        for i, text in enumerate(entries_label):
            tk.Label(result_frame, text=text).place(x=text_result_x, y=text_result_y+40*i)
            entry = tk.Entry(result_frame, width=entry_result_length)
            entry.place(x=entry_result_x, y=entry_result_y+40*i)
            entries[text] = entry

        for col in columns:
            table.heading(col, text=col)
            table.column(col, width=10)

        table.place(x=table_x, y=table_y, width=table_width-20, height=130)
        scrollbar = ttk.Scrollbar(mid_frame, orient=tk.VERTICAL, command=table.yview)
        table.configure(yscroll=scrollbar.set)
        scrollbar.place(x=table_x+table_width-20, y=table_y, height=230)

        controll_frame = tk.LabelFrame(mid_frame, text='Plot Hari tertentu', height=controll_frame_h, width=controll_frame_w)
        controll_frame.place(x=controll_frame_x, y=controll_frame_y)

        dropdown_var = StringVar()
        dropdown_picker = ttk.Combobox(controll_frame, textvariable=dropdown_var, state="readonly")
        dropdown_picker.place(x=40, y=10)

        def on_dropdown_change(event):
            selected_day = int(dropdown_var.get()) - 1
            create_plot(mid_frame, graph_1_x, graph_1_y, graph_1_w, graph_1_h, [section1[selected_day]], f'Plot hari ke-{selected_day + 1}')

        dropdown_picker.bind("<<ComboboxSelected>>", on_dropdown_change)

        create_plot(mid_frame, graph_1_x, graph_1_y, graph_1_w, graph_1_h)
        create_plot(mid_frame, graph_2_x, graph_2_y, graph_2_w, graph_2_h)

    def bot_frame(self):
        bot_frame = tk.Frame(self, background='#E0D8D8', height=int(0.07 * self.height), width=self.width)
        bot_frame.place(x=0, y=int(0.93 * self.height))

    def on_closing(self):
        self.quit()
        self.destroy()

if __name__ == "__main__":
    app = MeasurementApp()
    app.mainloop()
