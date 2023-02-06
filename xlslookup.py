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

####
import os
import win32api

def get_file_product_version(file_path):
    info = win32api.GetFileVersionInfo(file_path, "\\")
    ms = info['FileVersionMS']
    ls = info['FileVersionLS']
    return win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls), win32api.LOWORD(ls)

folder_path = "path/to/your/folder"
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    if os.path.isfile(file_path):
        major, minor, build, revision = get_file_product_version(file_path)
        product_version = "{}.{}.{}.{}".format(major, minor, build, revision)
        print("Product version for", filename, ":", product_version)
