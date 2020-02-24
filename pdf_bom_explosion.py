import datetime
import pandas as pd
import warnings
from openpyxl import load_workbook
from openpyxl.styles import Font
import re

import traverse_bom
import bin_locations

warnings.filterwarnings("ignore")


def supply_status(std, total_qty,
                  oh, ono, insp_oh,
                  p_type, due_date):
    today = pd.to_datetime(datetime.datetime.now().strftime('%Y-%m-%d'))
    due = pd.to_datetime(due_date).strftime('%Y-%m-%d')
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


# Explode Assembly
print('---------------exploding assembly---------------')
bom_explosion_df = traverse_bom.explode_bom(top_level='1016217', make_qty=1)
bom_explosion_df = bom_explosion_df.rename(columns={'PART NUMBER': 'Part',
                                                    'QTY': 'Comp Q/P'})

if not bom_explosion_df.empty:

    print('---------------getting supply status---------------')
    bom_expl_sum = bom_explosion_df.groupby('Part', as_index=False)['Comp Extd Qty']\
                                   .sum()\
                                   .rename(columns={'Comp Extd Qty': 'Total Needed'})

    bom_explosion_df = bom_explosion_df.merge(bom_expl_sum[['Part', 'Total Needed']], on='Part', how='left')

    bom_explosion_df['Supply Status'] = bom_explosion_df.apply(lambda x:
                                                               supply_status(x['Cost'], x['Total Needed'],
                                                                             x['OH'], x['ONO'],
                                                                             x['OH_Inspect'], x['TypeCode'],
                                                                             x['First Due']), axis=1)

    print('---------------getting floor locations---------------')
    # Get floor locations for components
    bin_df = bin_locations.get_bins(bom_explosion_df)

    # Add bin data to df
    bom_explosion_df = bom_explosion_df.merge(bin_df[['Part', 'Main Bins', 'Floor Bins']],
                                              on='Part', how='left')

    if traverse_bom.explode_bom.passed_args['ignore_epicor']:
        bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM'], axis=1)
    else:
        bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM', 'sub'], axis=1)

    print('---------------writing to excel---------------')
    # Create a Pandas Excel writer using openpyxl as the engine.
    assy_part = traverse_bom.explode_bom.passed_args['top_level']
    file_date = datetime.datetime.now().strftime("%m-%d-%Y")
    writer = pd.ExcelWriter(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion'
                            + '\\' + re.escape(assy_part) + f'_{file_date}.xlsx',
                            engine='openpyxl')
    book = writer.book
    bom_explosion_df.to_excel(writer, sheet_name='bom_explosion', index=False)

    # add hyperlink for print
    sheet = book['bom_explosion']
    part_list = bom_explosion_df['Part'].tolist()
    file_path = traverse_bom.explode_bom.passed_args['file_path']
    for i, row in enumerate(part_list):

        # sets hyperlink value
        sheet.cell(row=i+2, column=19).value = \
            f'=HYPERLINK("{file_path}{part_list[i]}.pdf","print")'

        # sets hyperlink format
        sheet.cell(row=i+2, column=19).font = Font(underline='single', color='0563C1')

    sheet.freeze_panes = 'A2'
    writer.save()

