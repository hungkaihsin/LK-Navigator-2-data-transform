import tkinter
import openpyxl
from openpyxl.styles import Font
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

# create the root window
root = tk.Tk()
root.title('Data Transform.exe')
root.resizable(False, False)
root.geometry('380x100')
root.iconbitmap('D:/Onedrive/OneDrive - Chinese Shining Missionary Association/02.Personal/02.Computer Related/05.Python/measure calculator/pythonProject/icon.ico')

# Label
lb_choosefile = ttk.Label(text='Choose file:', background = 'grey', foreground = 'white')
lb_choosefile.place(x = 0, y = 0)
lb_measuredata = ttk.Label(text = 'Measure data:', background = 'grey', foreground = 'white')
lb_measuredata.place(x = 0, y = 35)
lb_time = ttk.Label(text='Measure time:', background= 'grey', foreground= 'white')
lb_time.place(x = 200, y = 35)

# Entry
en_loadFile = ttk.Entry(width= 35)
en_loadFile.place(x = 85, y = 0)
en_measuredate = ttk.Entry(width= 8)
en_measuredate.place(x = 100, y = 35)
en_time = ttk.Entry(width= 8)
en_time.place(x= 300, y = 35)
en_filepath = ttk.Entry(width = 18)
en_filepath.place(x = 100, y = 73)

# Function
def load_file():
    filetypes = (
        ('Excel files', '*.xlsx'),('All files', '*.*')
    )
    if en_loadFile.get() is None:
        filepath = fd.askopenfilename(
        title = 'Open a file',
        initialdir = '/',
        filetypes = filetypes)
    else:
        filepath = fd.askopenfilename(
        title = 'Open a file',
        initialdir = '/',
        filetypes = filetypes)
        en_loadFile.delete(0,'end')
        en_loadFile.insert(0, filepath)

def Saveloacation():
    if en_filepath.get() is None:
        folderselect = fd.askdirectory()
        showinfo(title = 'Save', message = 'location saved')
    else:
        folderselect = fd.askdirectory()
        en_filepath.delete(0,'end')
        en_filepath.insert(0, folderselect)

def output():
    if len(en_loadFile.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please choose file')
    elif len(en_measuredate.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please type data')
    elif len(en_time.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please type time')
    elif len(en_filepath.get()) == 0:
        showinfo(title = 'Wrong', message = 'Please chose save location')
    else:
        file_route = en_loadFile.get()
        measuredata = en_measuredate.get()
        measuredata = int(measuredata)
        measuretime = en_time.get()
        measuretime = int(measuretime)
        time_period = float(measuretime / measuredata)
        wb_open = openpyxl.load_workbook(file_route)
        sheet = wb_open.worksheets[0]

    for transit in range(1, measuredata, 1):
        result = transit
        outcome = result * time_period
        sheet.cell(transit,4).value = outcome
    sheet.cell(1,2).value = "Data"
    sheet.cell(1,2).font = Font(bold = True)
    sheet.cell(1,5).value = "Time"
    sheet.cell(1,5).font = Font(bold = True)
    filesavepath = en_filepath.get()
    destination = filesavepath + "/Result.xlsx"
    wb_open.save(destination)
    showinfo(
        title = "Transform",
        message = "Done!",
    )



# Button
btn_loadfile = ttk.Button(text = '...',command = load_file, width = 4)
btn_loadfile.place(x = 340, y = 0)
btn_output = ttk.Button(text = 'Output', command = output)
btn_output.place(x = 250, y = 70)
btn_savelocation = ttk.Button(text = 'Save location', command = Saveloacation)
btn_savelocation.place(x = 0, y = 70)


# continue exist
root.mainloop()




'''
# create new Excel file, chose worksheet and save it

fn = 'Outcome.xlsx'
wb = openpyxl.Workbook()
sheet = wb.worksheets[0]

'''

