import os
import openpyxl

script_dir = os.path.dirname(__file__)
rel_path = "data/test.xlsx"
abs_file_path = os.path.join(script_dir, rel_path)

# Define variable to load the dataframe
dataframe = openpyxl.load_workbook(abs_file_path)

# Define variable to read sheet
dataframe1 = dataframe.active

# Iterate the read links in cells
for row in range(0, dataframe1.max_row):
    for col in dataframe1.iter_cols(1, dataframe1.max_column):
        link = col[row].value
        print(link)