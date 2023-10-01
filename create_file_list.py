import os
from openpyxl import Workbook
from itertools import zip_longest

# directory/folder path
dir_path_1 = input("path 1: ")
dir_path_2 = input("path 2: ")
# list to store files
res1 = []
res2 = []
# Iterate directory
for file_path in os.listdir(dir_path_1):
    # check if current file_path is a file
    if os.path.isfile(os.path.join(dir_path_1, file_path)):
        # add filename to list
        res1.append(file_path)
for file_path in os.listdir(dir_path_2):
    # check if current file_path is a file
    if os.path.isfile(os.path.join(dir_path_2, file_path)):
        # add filename to list
        res2.append(file_path)

file_list = zip_longest(res1, res2, fillvalue="")

wb = Workbook()
ws = wb.active


ws_write = wb.create_sheet()

for row in file_list:
    ws.append(row)

wb.save(filename='data.xlsx')
