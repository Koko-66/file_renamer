import os
import openpyxl


def get_path_to_matrix_file():
    """Get the file path from the user"""
    file_path = input("Path to the file with renaming details: ")
    return file_path


def extract_names_from_file(file_path):
    """Get filenames from Excel and create a dictionary"""
    wb_obj = openpyxl.load_workbook(file_path)

    # load the sheet
    sheet_obj = wb_obj.active
    # get number of rows for interation
    m_row = sheet_obj.max_row
    m_col = sheet_obj.max_column

    source_names = []
    target_names = []
    # loop over the rows and columns and crate two lsits
    for i in range(2, m_row + 1):
        # get cell object by row, column
        for j in range(1, m_col+1):
            cell_obj = sheet_obj.cell(row=i, column=j)
            if cell_obj.column == 1:
                source_names.append(cell_obj.value)
            else:
                target_names.append(cell_obj.value)
    names_dict = dict(zip(source_names, target_names))
    return names_dict


def rename_files(names_dict):
    """Rename files"""
    target_folder = input("Path to the files you want to rename: ")

    for file in os.listdir(target_folder):
        # check if file in the dictionary
        if file in names_dict.keys():
            old_name = os.path.join(target_folder, file)

            new_name = os.path.join(target_folder, names_dict[file])

            # rename the files
            os.rename(old_name, new_name)


def run():
    """Run program"""
    file_path = get_path_to_matrix_file()
    names_matrix = extract_names_from_file(file_path)
    rename_files(names_matrix)


if __name__ == '__main__':
    run()
