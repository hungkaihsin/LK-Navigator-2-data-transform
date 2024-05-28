import pandas as pd
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.alignment import Alignment
import matplotlib.pyplot as plt 

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
sheet.cell(1,2).font = Font(bold = True)
sheet["B1"].alignment = Alignment(horizontal= "center")
wb_open.save("test_excel.xlsx")

read_excel = pd.read_excel("test_excel.xlsx")
x_axis = read_excel["Time"]
y_axis = read_excel["Distance"]
plt.plot(x_axis, y_axis, color = "green")
plt.xlabel("Time(s)", fontsize = "12")
plt.ylabel("Distance(mm)", fontsize = "12")
plt.title("Actuator", fontsize = "15")
plt.ylim(-7,2)
plt.show()


