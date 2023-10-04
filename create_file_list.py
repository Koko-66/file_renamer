from itertools import zip_longest
import os
from openpyxl import Workbook
import pathlib


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
    # source = ["Source"]
    # target = ["Target"]
    # Iterate directory
    # for file_path in os.listdir(dir_source):
    #     # check if current file_path is a file
    #     if os.path.isfile(os.path.join(dir_source, file_path)):
    #         # add filename to list
    #         source.append(file_path)

    source = [file for root, dirs, files in os.walk(dir_source) for file in files if os.path.isfile(os.path.join(dir_source, file))]
    target = [os.path.join(root.split("\\")[-1], file) for root, dirs, files in os.walk(dir_target) for file in files]
    # for folder in os.listdir(dir_target):
    #    for file_path in os.listdir(os.path.join(dir_target, file_path)):
    #        print(f'{folder}_{file_path}')
        # check if current file_path is a file
        # if os.path.isfile(os.path.join(dir_target, file_path)):
            # add filename to list
    #    print(file_path)


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
