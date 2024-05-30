import openpyxl.workbook
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.alignment import Alignment
import matplotlib.pyplot as plt 
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import colorchooser as cc
from tkinter.messagebox import showinfo
import os

# Main root
root = tk.Tk()
root.title("Visualisation")
root.resizable(False, False)
root.geometry("400x400")
root.iconbitmap("logo.ico")


# Label
lb_choose_file = ttk.Label(text = "Choose file:")
lb_choose_file.place(x= 0, y = 3)
lb_measuredata = ttk.Label(text = "Measure data:")
lb_measuredata.place(x = 0, y = 37)
lb_measuredata_remind = ttk.Label(text = "(depend on amount data)")
lb_measuredata_remind.place(x=200, y = 37 )
lb_time = ttk.Label(text="Measure time:")
lb_time.place(x = 0, y = 77)
lb_time_remind = ttk.Label(text = "(s)")
lb_time_remind.place(x = 200, y= 77)
lb_color_choosen = tk.Label(height=1, width=2, bg= "green")
lb_color_choosen.place(x=195, y= 111)
lb_title_text = ttk.Label(text = "Title text:")
lb_title_text.place(x=0, y= 144)
lb_title_font_size = ttk.Label(text="Font size:")
lb_title_font_size.place(x=150, y=144)
lb_x_axis_title = ttk.Label(text="X axis title:")
lb_x_axis_title.place(x=0,y=177)
lb_y_axis_title = ttk.Label(text="Y axis title:")
lb_y_axis_title.place(x=0, y=210)
lb_x_axis_font_size = ttk.Label(text="Font size:")
lb_x_axis_font_size.place(x=160, y=177)
lb_y_axis_font_size = ttk.Label(text="Font size:")
lb_y_axis_font_size.place(x=160, y=210)



# Entry
en_choose_file = ttk.Entry(width= 25)
en_choose_file.place(x = 85, y = 0)
en_measuredate = ttk.Entry(width= 8)
en_measuredate.place(x = 100, y = 35)
en_time = ttk.Entry(width= 8)
en_time.place(x= 100, y = 75)
en_title_text = ttk.Entry(width= 8)
en_title_text.place(x=60, y=142)
en_title_font_size = ttk.Entry(width=8 )
en_title_font_size.place(x=210, y=142)
en_x_axis_title = ttk.Entry(width = 8)
en_x_axis_title.place(x=70, y=176)
en_y_axis_title = ttk.Entry(width=8)
en_y_axis_title.place(x=70, y=210)
en_x_axis_font_size = ttk.Entry(width=8)
en_x_axis_font_size.place(x=230, y=177)
en_y_axis_font_size = ttk.Entry(width=8)
en_y_axis_font_size.place(x=230, y=210)



# Function

# Color chooser
def colorchoose():
    color = cc.askcolor()
    color = color[1]
    print(str(color))
    lb_color_choosen.config(bg= color)


# Load File

def load_file():
    filetypes = (
        ('csv file', '*.csv'),('All files', '*.*')
    )
    if en_choose_file.get() is None:
        filepath = fd.askopenfilename(
        title = 'Open a file',
        initialdir = '/',
        filetypes = filetypes)
    else:
        filepath = fd.askopenfilename(
        title = 'Open a file',
        initialdir = '/',
        filetypes = filetypes)
        en_choose_file.delete(0,'end')
        en_choose_file.insert(0, filepath)

# Calculate and transform


def transit():
    if len(en_choose_file.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please choose file')
    elif len(en_measuredate.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please type data')
    elif len(en_time.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please type time')
    else:
        file_route = en_choose_file.get()
        measure_data = en_measuredate.get()
        measure_data = int(measure_data)
        measure_time = en_time.get()
        measure_time = int(measure_time)
        transition = float(measure_time / measure_data)

        load_csv_file = pd.read_csv(file_route, index_col= False, names=["Category", "Time","Distance"])
        load_csv_file.to_excel("Result.xlsx", index=False)
        wb_open = openpyxl.load_workbook("Result.xlsx")
        sheet = wb_open.worksheets[0]

        for transit in range(1, measure_data, 1):
            result = transit 
            outcome = result * transition
            sheet.cell((transit +1 ),2).value = outcome

        wb_open.save("Result.xlsx")
        excel_plot = pd.read_excel("Result.xlsx")
        x_axis = excel_plot["Time"]
        y_axis = excel_plot["Distance"]
        x_axis_title = en_x_axis_title.get()
        y_axis_title = en_y_axis_title.get()
        title = en_title_text.get()
        title_text_font_size = en_title_font_size.get()
        x_axis_font_size = en_x_axis_font_size.get()
        y_axis_font_size = en_y_axis_font_size.get()
        
        if len(en_title_font_size.get()) == 0:
            title_text_font_size = "12"
        
        if len(en_x_axis_font_size.get()) == 0:
            x_axis_font_size = "12"

        if len(en_y_axis_font_size.get()) == 0:
            y_axis_font_size = "12"


        
        plt.plot(x_axis, y_axis, color = lb_color_choosen["bg"])
        plt.xlabel(x_axis_title, fontsize = x_axis_font_size)
        plt.ylabel(y_axis_title, fontsize = y_axis_font_size)
        plt.title(title, fontsize = title_text_font_size)
        plt.ylim(-7,2)
        plt.show()
        
        os.remove("Result.xlsx")

     

# Button
btn_loadfile = ttk.Button(text = '...', command= load_file, width = 4)
btn_loadfile.place(x = 330, y = 0, width= 50)
btn_transit = ttk.Button(text = "Transit", command= transit)
btn_transit.place(x = 150, y = 250)
btn_choose_color = ttk.Button(text="Choose plot's line color", command= colorchoose)
btn_choose_color.place(x=0,y= 109)




root.mainloop()



