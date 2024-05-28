import pandas as pd
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import Alignment

# Read csv file
measure_data = pd.read_csv("Measure_data.csv", usecols=[2])

header_list = ["Distance"]

# Add Header
measure_data.to_excel("test_excel.xlsx", header= header_list, index= False)



wb_open = openpyxl.load_workbook("test_excel.xlsx")
sheet = wb_open.worksheets[0]
measure_time = 500
measure_data = 100000
transition = float(measure_time / measure_data)



for transit in range(1, measure_data, 1):
    count = transit
    outcome=  count * transition
    sheet.cell((transit + 1),2).value = outcome

sheet.cell(1,2).value = "Time"
sheet.cell(1,2).font = Font(bold = True, Alignment=(Alignment(horizontal="center")))
sheet.cell(1.2).alignment = Alignment




