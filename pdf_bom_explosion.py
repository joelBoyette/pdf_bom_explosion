import datetime

import traverse_bom
import bin_locations

import pandas as pd
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font
warnings.filterwarnings("ignore")


def supply_status(std, total_qty,
                  oh, ono, insp_oh,
                  p_type, due_date):
    today = pd.to_datetime(datetime.now().strftime('%Y-%m-%d'))
    due = pd.to_datetime(due_date.strftime('%Y-%m-%d'))
    if std == 0.0001:
        return 'Expense'
    elif total_qty == 0.0:
        return 'Consumed'
    elif oh >= total_qty:
        return 'Available'
    elif ono > total_qty and due < today:
        return 'Late'
    elif insp_oh + oh + ono >= total_qty and ono != 0.0:
        return 'ONO'
    elif p_type == 'M':
        return 'Mfg in house'
    elif insp_oh > 0:
        return 'Inspect'
    else:
        return 'Short'

# parts_df['Overall Status'] = parts_df.apply(lambda x: supply_status(x['Std'], x['Total Qty'],
#                                                                   x['sum_OH'], x['ONO'],
#                                                                   x['OH_Inspect'], x['Type'],
#                                                                   x['minDueDate']), axis=1)


# Explode Assembly
print('---------------exploding assembly---------------')
bom_explosion_df = traverse_bom.explode_bom(top_level='1021337', make_qty=1, ignore_epicor=False)
bom_explosion_df = bom_explosion_df.rename(columns={'PART NUMBER': 'Part',
                                                    'QTY': 'Comp Q/P'})

# Get floor locations for components
print('---------------getting floor locations---------------')
bin_df = bin_locations.get_bins(bom_explosion_df)

# Add bin data to df
bom_explosion_df = bom_explosion_df.merge(bin_df[['Part', 'Main Bins', 'Floor Bins']],
                                          on='Part', how='left')

bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM', 'sub'], axis=1)

print('---------------writing to excel---------------')
if not bom_explosion_df.empty:

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
    part_list = bom_explosion_df['Part'].tolist()
    file_path = traverse_bom.explode_bom.passed_args['file_path']
    for i, row in enumerate(part_list):

        # sets hyperlink value
        sheet.cell(row=i+2, column=17).value = \
            f'=HYPERLINK("{file_path}{part_list[i]}.pdf","print")'

        # sets hyperlink format
        sheet.cell(row=i+2, column=17).font = Font(underline='single', color='0563C1')

    sheet.freeze_panes = 'A2'
    writer.save()

