"""
Extracting data from multiple xls files and writing to a new one.

This module collects data from files (folder "calculations")
according to the list in the file "det_list _ *. Xlsx"
and combines this data into one file.
"""
from os.path import abspath
from openpyxl import load_workbook, Workbook
from def_xls import *

# get details list from xls file
date = '220401'
table_name = f'det_list_{date}.xlsx'
data_path = abspath(join('.', table_name))

wb = load_workbook(filename=data_path, data_only=True, read_only=True)
s_name = list(wb.sheetnames)[0]
ws_t = wb[s_name]

'''
# forming table header (make once for TABLE_HEADER)
table_header = get_table_header(ws_t)
'''

det_list = []
det_data = {}
for row in ws_t.iter_rows(min_row=2, min_col=1, max_row=ws_t.max_row,
                          max_col=ws_t.max_column, values_only=True):
    key_data = row[0]
    if key_data is not None:
        det_data[key_data] = [cell for cell in row[1::]]
        det_list.append(key_data)

wb.close()

# get file names according to the detail list
path_c = "\\calculations\\"
list_calc = os.listdir('.' + path_c)

det_list_files = {}
for ord_det in det_list:
    name_re_calc = str(ord_det + '_r\d{2}_\w{3,4}.xlsx')
    file_list_calc = file_name_re(name_re_calc, list_calc)
    det_list_files[ord_det] = file_list_calc

# get data from calculations
empty_list = []  # detail list without calculation
for det, file in det_list_files.items():
    calc_data = []
    if file:
        data_path = abspath(join('.' + '\\calculations\\', file))
        wb = load_workbook(filename=data_path, data_only=True, read_only=True)
        s_name = list(wb.sheetnames)
        for sn in s_name:
            if sn == 'матеріали':
                ws = wb[sn]
                '''
                # update table header (make once for TABLE_HEADER)
                table_header_2 = get_table_header(ws)
                table_header += table_header_2[2:5] + table_header_2[9:12]
                table_header.append('qty per det')
                '''
                # calc_data = []
                for row in ws.iter_rows(min_row=2, min_col=2, max_row=ws.max_row, max_col=12, values_only=True):
                    if row[0]:
                        r1 = row[1]  # check if part#
                        if not r1:
                            r1 = 1
                        qty = int(row[2]) * int(row[9])
                        calc_data_1 = [r1, row[2], row[3], row[8], row[9], row[10], qty]
                        calc_data.append(calc_data_1)
        wb.close()
    else:
        empty_list.append(det)

    det_data[det].append(calc_data)

# create and fill calculations data file
wb = Workbook()
ws = wb.active
ws.append(TABLE_HEADER)

for det in det_data:
    l = len(det_data[det][3])
    for i in range(0, l):
        row = [det]
        if i > 0:
            row = ['_']
        for r in det_data[det][0:3]:
            row.append(r)
        for c in det_data[det][3][i]:
            row.append(c)
        ws.append(row)

ws.column_dimensions['A'].width = 28
ws.column_dimensions['B'].width = 12
ws.column_dimensions['C'].width = 18
ws.column_dimensions['E'].width = 22
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 18
ws.column_dimensions['I'].width = 12
ws.column_dimensions['K'].width = 12

new_file = f'calculations_{date}.xlsx'
new_file_name = abspath(join('.', 'calculation_data', new_file))
wb.save(new_file_name)
wb.close()
