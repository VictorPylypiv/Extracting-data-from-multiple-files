"""
Extracting data from multiple xls files and writing to a new one.

This module collects data from files (folder "calculations")
according to the list in the file "det_list _ *. Xlsx"
and combines this data into one file.
"""
from def_xls import *
from hf_xls import get_data_calc_hf
from calc_to_xls import calculation_data


def create_calc_file(date1):
    path_c = "\\calculations\\"
    path_hf = "\\halffabricat_orders\\"
    sheet_name = "матеріали"

    det_data = imp_det_list(date1)
    det_list_files = get_files_name(det_data, path_c)
    det_data_c, empty_list = get_data_calc(det_data, det_list_files, path_c, sheet_name)
    det_data_c_hf = get_data_calc_hf(det_data_c, path_hf)
    calculation_data(det_data_c_hf, TABLE_HEADER, date, empty_list)


if __name__ == "__main__":
    date = input("Enter the date of the detail group: ")
    # date = '2022-04-01'
    create_calc_file(date)
