import pandas as pd

# Read csv file
measure_data = pd.read_csv("Measure_data.csv")
print(measure_data.head())

header_list = ["Categroy", "Time", "Distance"]

# Add Header
measure_data.to_csv("Measure_data2.csv", header= header_list, index= False)


usecols = ["Distance"]

measure_data2 = pd.read_csv("Measure_data2.csv", index_col=0, usecols=usecols)
print(measure_data2.head())



measure_time = 500
measure_data = 100000
transition = float(measure_time / measure_data)

print(transition)

for transit in range(1, measure_data, 1):
    count = transit
    calculate =  count * transition
    

    