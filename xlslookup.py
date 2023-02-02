import openpyxl
import os

# Load the workbook
workbook = openpyxl.load_workbook('file.xlsx')

# Select the sheet
sheet = workbook['Sheet1']

# Specify the column to be read (in this case, column 'A')
column = sheet['A']

# Store the values of each cell in the column in a list
column_values = [cell.value.split(" ")[0] for cell in column]

# Define a function to parse the directory recursively
def search_files(path):
    for root, dirs, files in os.walk(path):
        for file in files:
            if file in column_values:
                print(f"File '{file}' is present in the directory '{root}'")
                break
    else:
        print(f"File '{file}' is not present in the directory")

# Call the function to search the files
search_files("/path/to/directory")
