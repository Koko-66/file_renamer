from itertools import zip_longest
import os
from openpyxl import Workbook


def get_paths():
    """Get directory paths"""
    # directory/folder path
    dir_source = input(
        "Path to the folder with the source files: ")
    dir_target = input("Path to the folder with the target files: ")
    # list to store files
    return dir_source, dir_target


def get_file_names():
    """Extract file names from folders"""
    dir_source, dir_target = get_paths()

    source = [file for root, dirs, files in os.walk(dir_source) for file in sorted(
        files, key=lambda file: len(file)) if os.path.isfile(os.path.join(dir_source, file))]
    target = [os.path.join(root.split("\\")[-1], file) for root, dirs, files in os.walk(
        dir_target) for file in sorted(files, key=lambda file: len(file))]

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
