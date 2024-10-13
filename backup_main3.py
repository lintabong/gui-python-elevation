import time
import ctypes
import locale
import tkinter
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from helper import process_data
from constant import *

ctypes.windll.shcore.SetProcessDpiAwareness(2)
locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')

class MeasurementApp(tkinter.Tk):
    def __init__(self):
        super().__init__()
        self.title('Measurement App')

        self.dpi = self.winfo_fpixels('1i')

        self.width = int(1080 * (self.dpi / 96))
        self.height = int(700 * (self.dpi / 96))
        x_pos = int((self.winfo_screenwidth() / 2) - (self.width / 2))
        y_pos = int((self.winfo_screenheight() / 2) - (self.height / 2))

        self.geometry(f'{self.width}x{self.height}+{x_pos}+{y_pos}')
        
        self.file_path = None

        self.height_proportion = [0.1, 0.83, 0.07]

        self.top_frame()
        self.mid_frame()
        self.bot_frame()
        
        self.protocol('WM_DELETE_WINDOW', self.on_closing)

    def top_frame(self):
        frame = tkinter.Frame(self, background='#292F36', height=int(self.height_proportion[0] * self.height), width=self.width)
        frame.place(x=0, y=0)

        image = Image.open('logo.png')
        image = image.resize((90, 60))
        logo = ImageTk.PhotoImage(image)

        label_logo = tkinter.Label(frame, image=logo, bg='#292F36')
        label_logo.image = logo
        label_logo.place(x=10, y=10)

        label_time = tkinter.Label(frame, text="", fg='white', bg='#292F36', font=('Helvetica', 16))
        label_time.place(x=self.width-130, y=30)

        def update_time():
            current_time = time.strftime('%H:%M:%S')
            label_time.config(text=current_time)
            label_time.after(1000, update_time)

        update_time()

    def mid_frame(self):
        frame = tkinter.Frame(self, background='#FAF5F1', height=int(self.height_proportion[1] * self.height), width=self.width)
        frame.place(x=0, y=int(self.height_proportion[0] * self.height))

        def create_empty_graph(frame, x, y, width, height):
            fig, ax = plt.subplots(figsize=(3, 2))
            ax.plot([], [])
            ax.set_xlim(0, 10)
            ax.set_ylim(0, 10)
            ax.set_xlabel('X-axis')
            ax.set_ylabel('Y-axis')

            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().place(x=x, y=y, width=width, height=height)

        create_empty_graph(frame, GRAPH1_X, GRAPH1_Y, GRAPH1_W, GRAPH1_H)
        create_empty_graph(frame, GRAPH2_X, GRAPH2_Y, GRAPH2_W, GRAPH2_H)

    def bot_frame(self):
        frame = tkinter.Frame(self, background='#E0D8D8', height=int(self.height_proportion[2] * self.height), width=self.width)
        frame.place(x=0, y=int((self.height_proportion[0] + self.height_proportion[1]) * self.height))
        tkinter.Label(frame, text=COPYRIGHT_TEXT, bg='#E0D8D8').place(x=self.width-len(COPYRIGHT_TEXT)*10, y=20)

    def on_closing(self):
        self.quit()
        self.destroy()

if __name__ == "__main__":
    app = MeasurementApp()
    app.mainloop()
