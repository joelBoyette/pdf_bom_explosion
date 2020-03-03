from openpyxl.styles import Font
import numpy as np
import tkinter as tk
import re
import datetime
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

import logging
from log_config import log_location, log_filemode, log_format, log_datefmt
logger = logging.getLogger(__name__)
logging.basicConfig(filename=log_location + 'app.log',
                    filemode=log_filemode,
                    format=log_format,
                    datefmt=log_datefmt,
                    level=logging.INFO)
logging.getLogger('pdfdocument').setLevel(logging.CRITICAL)
logging.getLogger('pdfinterp').setLevel(logging.CRITICAL)

# pyinstaller to make a distribution file for users to use program
# https://pyinstaller.readthedocs.io/en/stable/usage.html


def supply_status(std, total_qty,
                  oh, ono, insp_oh,
                  p_type, due_date):
    today = pd.to_datetime(datetime.datetime.now()).strftime('%Y-%m-%d')
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


def main():

    window = tk.Tk()
    window.geometry('350x200')

    top_level_lbl = tk.Label(window, text='Top Level', padx=20, pady=10)
    top_level_lbl.grid(column=0, row=0)

    top_level_txtbx = tk.Entry(window, width=10)
    top_level_txtbx.grid(column=0, row=1, padx=1, pady=1)

    qty_lbl = tk.Label(window, text='Make Qty', padx=20, pady=10)
    qty_lbl.grid(column=0, row=2)

    qty_txtbx = tk.Entry(window, width=10)
    qty_txtbx.grid(column=0, row=3, padx=1, pady=1)

    log_lbl = tk.Label(window, text='Log')
    log_lbl.grid(column=0, row=6)

    def explode_assembly():


        # user input
        top_level_input = top_level_txtbx.get()
        qty_input = qty_txtbx.get()

        # Explode Assembly
        print('---------------retrieving epicor data---------------')
        log_lbl.config(text='---------------retrieving epicor data---------------')
        log_lbl.update()

        import traverse_bom

        print('---------------exploding assembly---------------')
        log_lbl.config(text='---------------exploding assembly---------------')
        log_lbl.update()

        bom_explosion_df = traverse_bom.explode_bom(top_level=top_level_input, make_qty=int(qty_input))
        bom_explosion_df = bom_explosion_df.rename(columns={'PART NUMBER': 'Part',
                                                            'QTY': 'Comp Q/P'})

        if not bom_explosion_df.empty:

            print('---------------getting supply status---------------')
            bom_expl_sum = bom_explosion_df.groupby('Part', as_index=False)['Comp Extd Qty'] \
                .sum() \
                .rename(columns={'Comp Extd Qty': 'Total Needed'})

            bom_explosion_df = bom_explosion_df.merge(bom_expl_sum[['Part', 'Total Needed']], on='Part', how='left')

            bom_explosion_df['Supply Status'] = bom_explosion_df.apply(lambda x:
                                                                       supply_status(x['Cost'], x['Total Needed'],
                                                                                     x['OH'], x['ONO'],
                                                                                     x['OH_Inspect'], x['TypeCode'],
                                                                                     x['First Due']), axis=1)

            print('---------------getting floor locations---------------')
            import bin_locations
            # Get floor locations for components
            bin_df = bin_locations.get_bins(bom_explosion_df)

            # Add bin data to df
            bom_explosion_df = bom_explosion_df.merge(bin_df[['Part', 'Main Bins', 'Floor Bins']],
                                                      on='Part', how='left')

            if traverse_bom.explode_bom.passed_args['ignore_epicor']:
                bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM'], axis=1)
            else:
                bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM', 'sub'], axis=1)

            print('---------------Formatting and writing to excel---------------')
            bom_explosion_df['Top Level'] = bom_explosion_df['Sort Path'][0][1:]
            bom_explosion_df['BOM Details [BOM Path | Parent Q/P | Comp Q/P]'] = \
                bom_explosion_df[['Sort Path', 'Assembly Q/P', 'Comp Q/P']].astype(str) \
                    .apply(lambda x: '[' + '-'.join(x) + ']' + '\n' if x.all != '' else set(x), axis=1)

            bom_explosion_df = bom_explosion_df.rename(columns={'PartDescription': 'Description',
                                                                'TypeCode': 'Type'})

            bom_explosion_df['Dwg Link'] = ''
            bom_explosion_df = bom_explosion_df[
                ['Part', 'Description', 'Type', 'Dwg Link', 'Total Needed', 'Supply Status',
                 'Main Bins', 'Floor Bins', 'OH', 'OH_Inspect', 'ONO', 'First Due', 'DMD',
                 'Buyer', 'Cost', 'Top Level', '# Top Level to Make',
                 'BOM Details [BOM Path | Parent Q/P | Comp Q/P]']]
            bom_explosion_df = bom_explosion_df.astype(str)

            # https://thispointer.com/python-how-to-use-if-else-elif-in-lambda-functions/
            pivot_f = lambda x: x.sum() if np.issubdtype(x.dtype, np.number) \
                else (''.join(list(set(x)))
                      if x.name in ['BOM Details [BOM Path | Parent Q/P | Comp Q/P]']
                      else list(set(x))[0])

            cols = bom_explosion_df.columns[~bom_explosion_df.columns.isin(['Part'])].tolist()
            pivtbl = pd.pivot_table(bom_explosion_df,
                                    index=['Part'],
                                    values=cols,
                                    aggfunc=pivot_f,
                                    fill_value='').reindex(columns=cols).reset_index()

            # Create a Pandas Excel writer using openpyxl as the engine.
            assy_part = bom_explosion_df['Top Level'][0]
            file_date = datetime.datetime.now().strftime("%m-%d-%Y")
            writer = pd.ExcelWriter(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion'
                                    + '\\' + re.escape(assy_part) + f'_{file_date}.xlsx',
                                    engine='openpyxl')
            book = writer.book
            pivtbl.to_excel(writer, sheet_name='bom_explosion', index=False)

            # add hyperlink for print
            sheet = book['bom_explosion']
            part_list = pivtbl['Part'].tolist()
            file_path = traverse_bom.explode_bom.passed_args['file_path']
            for i, row in enumerate(part_list):
                # sets hyperlink value
                sheet.cell(row=i + 2, column=4).value = \
                    f'=HYPERLINK("{file_path}{part_list[i]}.pdf","print")'

                # sets hyperlink format
                sheet.cell(row=i + 2, column=4).font = Font(underline='single', color='0563C1')

            sheet.freeze_panes = 'A2'
            writer.save()

    btn = tk.Button(window, text="Explode BOM", padx=1, pady=1, command=explode_assembly)
    btn.grid(column=0, row=5)

    window.mainloop()


if __name__ == '__main__':
    main()
