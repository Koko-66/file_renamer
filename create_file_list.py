from itertools import zip_longest
import os
from openpyxl import Workbook


def get_paths():
    """Get directory paths"""
    # directory/folder path
    dir_source = input(
        "Path to the folder with the files you need to rename: ")
    dir_target = input("Path to the folder with the target files: ")
    # list to store files
    return dir_source, dir_target


def get_file_names():
    """Extract file names from folders"""
    dir_source, dir_target = get_paths()
    source = ["Source"]
    target = ["Target"]
    # Iterate directory
    for file_path in os.listdir(dir_source):
        # check if current file_path is a file
        if os.path.isfile(os.path.join(dir_source, file_path)):
            # add filename to list
            source.append(file_path)
    for file_path in os.listdir(dir_target):
        # check if current file_path is a file
        if os.path.isfile(os.path.join(dir_target, file_path)):
            # add filename to list
            target.append(file_path)

    filename_data = zip_longest(source, target, fillvalue="")
    return filename_data


def save_file_list_to_excel(file_list):
    """Save filenames to Excel file"""
    wb = Workbook()
    ws = wb.active
    wb.create_sheet()

    for row in file_list:
        ws.append(row)

    wb.save(filename='data.xlsx')


if __name__ == "__main__":
    file_list = get_file_names()
    save_file_list_to_excel(file_list)
