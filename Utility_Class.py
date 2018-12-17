import os
from xlsxwriter.workbook import Workbook



def get_filenames(data_path):  # Public method definition
    '''
    :param data_path: The path of the excel file to be processed
    :return: All excel names in this path. Data type is a list
    '''
    filenames = []
    for i in os.walk(data_path):
        for filename in i[-1]:
            full_filename = os.path.join(i[0], filename)
            filenames.append(full_filename)
    return filenames


def list_to_excel(end_xls, sheet_name, list_data):
    workbook = Workbook(end_xls)
    worksheet = workbook.add_worksheet(sheet_name)  # add a sheet
    data = list_data
    for row_num, row_data in enumerate(data):
        worksheet.write_row(row_num + 1, 0, row_data)
    workbook.close()

def isInt(num):
    # Whether the data is of integer type
    try:
        num = int(str(num))
        return isinstance(num, int)
    except:
        return False


def string_split(string, start_location, end_location=None):
    # Start with the start_location
    return string[start_location: end_location]


def list_split(list, start_location, end_location=None):
    # Start with the start_location, end with end_location-1
    return list[start_location: end_location]
