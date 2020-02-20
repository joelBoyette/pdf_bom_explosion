
import traverse_bom

import pandas as pd
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font
warnings.filterwarnings("ignore")

# run recursion
bom_explosion_df = traverse_bom.explode_bom(top_level='1021045', make_qty=1)

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
file_path = traverse_bom.explode_bom.passed_args['file_path']
for i, row in enumerate(part_list):

    # sets hyperlink value
    sheet.cell(row=i+2, column=17).value = \
        f'=HYPERLINK("{file_path}{part_list[i]}.pdf","print")'

    # sets hyperlink format
    sheet.cell(row=i+2, column=17).font = Font(underline='single', color='0563C1')

sheet.freeze_panes = 'A2'
writer.save()

