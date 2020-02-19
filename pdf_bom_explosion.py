import camelot
import pandas as pd
import numpy as np
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font
warnings.filterwarnings("ignore")

from pdf_bom import read_pdf_bom


epicor_data = \
    pd.read_excel(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\test_epicor_data.xlsx',
                  sheet_name='data')

# ################ CAMELOT ####################
# https://camelot-py.readthedocs.io/en/master/user/advanced.html
def clean_pdf_camelot_scale15(part, file_path):
    full_path = file_path + f'{part}.pdf'
    find_column_dict = {}
    # https://camelot-py.readthedocs.io/en/master/user/advanced.html
    pdf_df = camelot.read_pdf(full_path, line_scale=15, shift_text=[''])[0].df
    pdf_df = pdf_df.replace('\\n', '', regex=True)
    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]
    df_final = df_final.dropna(how='all')
    df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'})

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(
        '(^[.])|([\/])', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]

    df_final['FUNCTION'] = 'camelot_scale15'

    return df_final


def traverse_bom(top_level, qty, file_path):

    traverse_bom.part = top_level
    traverse_bom.assy_qty = qty
    traverse_bom.sort_path += '\\' + traverse_bom.part

    print(f'traverse bom part {traverse_bom.part} level {traverse_bom.level} path {traverse_bom.sort_path}')

    # Get bom from pdf and merge epicor data
    df = read_pdf_bom(traverse_bom.part, file_path)
    df = df.merge(epicor_data, on='PART NUMBER', how='left')

    # put non phantoms in df
    df_no_phantom = df.loc[~df['Phantom BOM']]
    df_no_phantom['Assembly'] = traverse_bom.part
    df_no_phantom['Assembly Q/P'] = traverse_bom.assy_qp
    df_no_phantom['Assembly Extd Qty'] = traverse_bom.assy_qty
    df_no_phantom['Comp Extd Qty'] = df_no_phantom['Assembly Extd Qty'] * df_no_phantom['QTY'].astype(int)
    df_no_phantom['# Top Level to Make'] = df_no_phantom['Assembly Extd Qty'] / df_no_phantom['Assembly Q/P']
    df_no_phantom['Level'] = traverse_bom.level
    df_no_phantom['Sort Path'] = traverse_bom.sort_path
    df_no_phantom['Dwg Link'] = ''

    # append non phantoms to final list
    traverse_bom.df_final = traverse_bom.df_final.append(df_no_phantom)

    # capture the sort path before recursion.  used to set the sort path on phantoms
    sort_path_reset = traverse_bom.sort_path

    df_phantoms = df.loc[df['Phantom BOM']][['PART NUMBER', 'QTY']]

    if not df_phantoms.empty:

        # slice phantoms and traverse to get boms
        df_phantom_dict = df_phantoms.to_dict("list")

        for index, phantom in enumerate(df_phantom_dict):

            traverse_bom.level += 1

            phantom = df_phantom_dict["PART NUMBER"][index]
            phantom_qty = int(df_phantom_dict["QTY"][index])
            phantom_extd_qty = phantom_qty * traverse_bom.assy_qty

            traverse_bom.assy_qp = phantom_qty

            # recursion
            traverse_bom(phantom, phantom_extd_qty, file_path)

            # reset bom level and sort path for correct leveling
            traverse_bom.level -= 1
            traverse_bom.sort_path = sort_path_reset

    return traverse_bom.df_final


# set recursion variables
traverse_bom.df_final = pd.DataFrame()
traverse_bom.level = 0
traverse_bom.sort_path = ''
traverse_bom.assy_qty = 0
traverse_bom.assy_qp = 1

base_path = r'\\vimage\latest' + '\\'

# run bom recursion
bom_explosion_df = traverse_bom('1021045', 5, base_path)


# Load and clear old data
excel_book = r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\output.xlsx'
book = load_workbook(excel_book)
book['Sheet1'].sheet_state = 'visible'

keep_sheets = ['Sheet1']
for sheetName in book.sheetnames:
    if sheetName not in keep_sheets:
        del book[sheetName]
book.save(excel_book)

#  Create a Pandas Excel writer using openpyxl as the engine.
writer = pd.ExcelWriter(excel_book, engine='openpyxl')
writer.book = book

# write to file
bom_explosion_df.to_excel(writer, sheet_name='bom_explosion', index=False)

# add hyperlink for print
sheet = book['bom_explosion']
part_list = bom_explosion_df['PART NUMBER'].tolist()
for i, row in enumerate(part_list):

    # sets hyperlink value
    sheet.cell(row=i+2, column=16).value = \
        f'=HYPERLINK("{base_path}{part_list[i]}.pdf","print")'

    # sets hyperlink format
    sheet.cell(row=i+2, column=16).font = Font(underline='single', color='0563C1')

sheet.freeze_panes = 'A2'
writer.save()

