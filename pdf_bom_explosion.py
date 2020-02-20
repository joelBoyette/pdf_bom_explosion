
import pandas as pd
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font
warnings.filterwarnings("ignore")

import pdf_bom


epicor_data = \
    pd.read_excel(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\test_epicor_data.xlsx',
                  sheet_name='data')

subassy_df = pd.read_excel(r'\\Vfile\d\MaterialPlanning\MPPublic\manufactured_part_routings.xlsm')

subassy_df = subassy_df.loc[(subassy_df['ResourceGrpID'] != 'FAB')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'ELEC')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'HOSE')
                            ]
subassy_df = subassy_df['MtlPartNum'].unique()


def traverse_bom(top_level, make_qty, file_path=r'\\vimage\latest' + '\\'):

    traverse_bom.passed_args = locals()

    traverse_bom.part = top_level
    traverse_bom.assy_qty = make_qty
    traverse_bom.sort_path += '\\' + traverse_bom.part

    print(f'traverse bom part {traverse_bom.part} '
          f'level {traverse_bom.level} path {traverse_bom.sort_path} '
          f'assy qty {traverse_bom.assy_qty}')

    # Get bom from pdf and merge epicor data
    df = pdf_bom.read_pdf_bom(traverse_bom.part, file_path)
    df = df.merge(epicor_data, on='PART NUMBER', how='left')
    df.loc[df['PART NUMBER'].isin(subassy_df), 'sub'] = 'yes'
    df.loc[~df['PART NUMBER'].isin(subassy_df), 'sub'] = 'no'

    # put non phantoms in df
    df_no_phantom = df.loc[(~df['Phantom BOM']) & (df['sub'] == 'no')]
    df_no_phantom['# Top Level to Make'] = make_qty / traverse_bom.assy_qp
    df_no_phantom['Assembly'] = traverse_bom.part
    df_no_phantom['Assembly Q/P'] = traverse_bom.assy_qp
    df_no_phantom['Assembly Make Qty'] = df_no_phantom['# Top Level to Make'] * df_no_phantom['Assembly Q/P']
    df_no_phantom['Comp Extd Qty'] = df_no_phantom['Assembly Make Qty'] * df_no_phantom['QTY'].astype(int)

    df_no_phantom['Level'] = traverse_bom.level
    df_no_phantom['Sort Path'] = traverse_bom.sort_path
    df_no_phantom['Dwg Link'] = ''

    # append non phantoms to final list
    traverse_bom.df_final = traverse_bom.df_final.append(df_no_phantom)

    # capture the sort path before recursion.  used to set the sort path on phantoms
    sort_path_reset = traverse_bom.sort_path

    # check for phantoms and traverse boms
    df_phantoms = df.loc[(df['Phantom BOM']) | (df['sub'] == 'yes')][['PART NUMBER', 'QTY']]

    if not df_phantoms.empty:

        for idx, df_row in df_phantoms.iterrows():

            traverse_bom.level += 1

            phantom = df_phantoms.loc[idx, 'PART NUMBER']
            phantom_qty = int(df_phantoms.loc[idx, 'QTY'])
            phantom_extd_qty = traverse_bom.assy_qty * phantom_qty
            traverse_bom.assy_qp = phantom_qty

            # recursion through lower levels
            traverse_bom(phantom, phantom_extd_qty, file_path)

            # reset bom level and sort path for next traversal
            traverse_bom.level -= 1
            traverse_bom.sort_path = sort_path_reset
            traverse_bom.assy_qty = make_qty

    return traverse_bom.df_final


# setup recursion variables
traverse_bom.df_final = pd.DataFrame()
traverse_bom.level = 0
traverse_bom.sort_path = ''
traverse_bom.assy_qty = 0
traverse_bom.assy_qp = 1

# run recursion
bom_explosion_df = traverse_bom(top_level='1021045', make_qty=1)

# Load xlsx and clear old data
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
bom_explosion_df.to_excel(writer, sheet_name='bom_explosion', index=False)

# add hyperlink for print
sheet = book['bom_explosion']
part_list = bom_explosion_df['PART NUMBER'].tolist()
file_path = traverse_bom.passed_args['file_path']
for i, row in enumerate(part_list):

    # sets hyperlink value
    sheet.cell(row=i+2, column=17).value = \
        f'=HYPERLINK("{file_path}{part_list[i]}.pdf","print")'

    # sets hyperlink format
    sheet.cell(row=i+2, column=17).font = Font(underline='single', color='0563C1')

sheet.freeze_panes = 'A2'
writer.save()

