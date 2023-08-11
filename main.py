import pandas as pd
import os
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, colors, Border, Side, Alignment, PatternFill


def generate_sheets(list_):
    wb = Workbook()
    ws = wb.active
    ws.title = str(list_[0])
    for i in list_[1:]:
        wb.create_sheet(str(i))
    wb.save('sheets.xlsx')


def write(list_, list_2, str_):
    for i in range(len(list_)):
        wb = load_workbook(path + r'\sheets.xlsx')
        ws = wb[str(list_[i])]
        ws.cell(row=1, column=1, value=str_+'-'+list_[i])
        ws.merge_cells(range_string='A1:M1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
#        wb.save('sheets.xlsx')
        print(df_structure[df_structure['机架编号'] == list_2[i]])


if __name__ == '__main__':
    path = os.getcwd()
    df_structure = pd.read_excel(path + r'\structure.xlsx')
    room_name = df_structure.columns[-1]
    list_shelf = list({}.fromkeys(df_structure['机架编号'].to_list()).keys())
    list_shelf_name = [item + '端截面图' for item in list_shelf]
    generate_sheets(list_shelf_name)
    write(list_shelf_name, list_shelf, room_name)



