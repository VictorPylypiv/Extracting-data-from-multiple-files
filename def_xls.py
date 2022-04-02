import os
import re
from os.path import join

TABLE_HEADER = ['ord_det', 'ord', 'det', 'qty', 'part# ', 'qty of p.',
                'material name', 'material type', 'material per 1',
                'unit', 'qty per det']

# get file list var 1:
path_dir_c = join('.', 'calculations')


def file_list(directory):
    filelist = []
    for root, dirs, files in os.walk(directory):
        for filename in files:
            filelist.append(filename)
    return filelist


fl = file_list(path_dir_c)

# get file list var 2:
path_c = "\\calculations\\"
list_calc = os.listdir('.' + path_c)


# get list of files in detail list version 1:
def file_name_re_1(name_re, file_list):
    name_f = None
    for name in file_list:
        name_m = re.search(name_re, name)
        if name_m:
            name_f = name_m[0]
    return name_f


# get list of files in detail list version 2:
def file_name_re(name_re, file_list):
    name_comp = re.compile(name_re)
    for name in file_list:
        name_m = name_comp.match(name)
        if name_m:
            return name_m.group()


# get table headers
def get_table_header(ws, min_column=1, max_column=None):
    if not max_column:
        max_column = ws.max_column
    header = [cell.value for cell in next(ws.iter_rows(
        min_row=1, min_col=min_column, max_row=1, max_col=max_column))]
    return header
