import win32com
from win32com.client import Dispatch
import pdf_bom
import pandas as pd
import logging
import datetime
import warnings
import numpy as np
from openpyxl.styles import Font
import re


warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

print('---------------retrieving epicor data---------------')

logger.critical('---------------retrieving epicor data---------------')
epicor_data = pd.read_excel(r'\\vfile\MPPublic\ECN Status\ecn_data.xlsm',
                            sheet_name='epicor_part_data')

min_due = pd.read_excel(r'\\vfile\MPPublic\pdf_bom_explosion\mindue_inspOH_data.xlsm',
                        sheet_name='min_due')

inspect_oh = pd.read_excel(r'\\vfile\MPPublic\pdf_bom_explosion\mindue_inspOH_data.xlsm',
                           sheet_name='inspect_oh')
inspect_oh['PartNum'] = inspect_oh['PartNum'].astype(str)

# add min due date and inspection OH
epicor_data = epicor_data.merge(min_due[['PartNum', 'First Due']], on='PartNum', how='left')
epicor_data = epicor_data.merge(inspect_oh[['PartNum', 'SumOfOurQty']], on='PartNum', how='left')
epicor_data = epicor_data.rename(columns={'SumOfOurQty': 'OH_Inspect'})

subassy_df = pd.read_excel(r'\\Vfile\d\MaterialPlanning\MPPublic\manufactured_part_routings.xlsm')
subassy_df = subassy_df.loc[(subassy_df['ResourceGrpID'] != 'FAB')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'ELEC')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'HOSE')
                            & (subassy_df['ResourceGrpID'] != 'PAINT')]

subassy_df = subassy_df['MtlPartNum'].unique()


# recursion through the BOM
def explode_bom(top_level, make_qty, file_path, ignore_epicor):

    # locals function provides list of arguments passed for a function
    explode_bom.passed_args = locals()

    explode_bom.part = top_level
    explode_bom.assy_qty = make_qty
    explode_bom.sort_path += '\\' + explode_bom.part

    logging_information = f'traverse bom part {explode_bom.part} '\
                          f'level {explode_bom.level} path {explode_bom.sort_path} '\
                          f'assy qty {explode_bom.assy_qty}'

    print(logging_information)
    logger.critical(logging_information)

    # Get bom from pdf and merge epicor data
    df = pdf_bom.read_pdf_bom(explode_bom.part, file_path)

    if df.empty:
        pass
    else:

        # use epicor flags to determine which parts to explode
        if not ignore_epicor:

            df = df.merge(epicor_data[['PartNum', 'PartDescription', 'PhantomBOM', 'TypeCode',
                                       'Cost', 'OH', 'ONO', 'DMD', 'Buyer', 'First Due', 'OH_Inspect']],
                          left_on='PART NUMBER', right_on='PartNum', how='left').fillna(0)

            df = df.drop('PartNum', axis=1)
            df.loc[df['PART NUMBER'].isin(subassy_df), 'sub'] = 'yes'
            df.loc[~df['PART NUMBER'].isin(subassy_df), 'sub'] = 'no'

            # put non phantoms in df
            df_no_explode = df.loc[(df['PhantomBOM'] == 0) & (df['sub'] == 'no')]

            df_no_explode['# Top Level to Make'] = make_qty / explode_bom.assy_qp
            df_no_explode['Assembly'] = explode_bom.part
            df_no_explode['Assembly Q/P'] = explode_bom.assy_qp
            df_no_explode['Assembly Make Qty'] = df_no_explode['# Top Level to Make'] * df_no_explode['Assembly Q/P']
            df_no_explode['Comp Extd Qty'] = df_no_explode['Assembly Make Qty'] * df_no_explode['QTY'].astype(int)
            df_no_explode['Level'] = explode_bom.level
            df_no_explode['Sort Path'] = explode_bom.sort_path
            df_no_explode['Dwg Link'] = ''

            # append non phantoms to final list
            explode_bom.df_final = explode_bom.df_final.append(df_no_explode)

            # capture the sort path before recursion.  used to reset the sort path after phantom is traversed
            sort_path_reset = explode_bom.sort_path

            # check for phantoms and traverse boms
            df_explode = df.loc[(df['PhantomBOM'] == 1) | (df['sub'] == 'yes')][['PART NUMBER', 'QTY']]

        else:
            # do not care about epicor flags, keep exploding as long as prints have BOM's

            df = df.merge(epicor_data[['PartNum', 'PartDescription', 'PhantomBOM', 'TypeCode',
                                       'Cost', 'OH', 'ONO', 'DMD', 'Buyer', 'First Due', 'OH_Inspect']],
                          left_on='PART NUMBER', right_on='PartNum', how='left').fillna(0)
            df = df.drop('PartNum', axis=1)

            df_no_explode = df
            df_no_explode['# Top Level to Make'] = make_qty / explode_bom.assy_qp
            df_no_explode['Assembly'] = explode_bom.part
            df_no_explode['Assembly Q/P'] = explode_bom.assy_qp
            df_no_explode['Assembly Make Qty'] = df_no_explode['# Top Level to Make'] * df_no_explode['Assembly Q/P']
            df_no_explode['Comp Extd Qty'] = df_no_explode['Assembly Make Qty'] * df_no_explode['QTY'].astype(int)

            df_no_explode['Level'] = explode_bom.level
            df_no_explode['Sort Path'] = explode_bom.sort_path
            df_no_explode['Dwg Link'] = ''

            # append non phantoms to final list
            explode_bom.df_final = explode_bom.df_final.append(df_no_explode)

            # capture the sort path before recursion.  used to reset the sort path after phantom is traversed
            sort_path_reset = explode_bom.sort_path

            # there will never be phantoms due to ignoring flags so make df_explode equal the original df
            df_explode = df

        if not df_explode.empty:

            for df_idx, df_row in df_explode.iterrows():

                explode_bom.level += 1
                explode_part = df_explode.loc[df_idx, 'PART NUMBER']
                explode_qty = int(df_explode.loc[df_idx, 'QTY'])
                explode_extd_qty = explode_bom.assy_qty * explode_qty
                explode_bom.assy_qp = explode_qty

                # recursion through lower levels
                explode_bom(explode_part, explode_extd_qty, file_path, ignore_epicor)

                # reset bom level and sort path for next traversal
                explode_bom.level -= 1
                explode_bom.sort_path = sort_path_reset
                explode_bom.assy_qty = make_qty

    return explode_bom.df_final


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


# Explode Assembly
def explode_assembly(top_level_input, qty_input, file_path, ignore_epicor, email_supp, supp_adr):

    file_path += '\\'

    print('---------------exploding assembly---------------')
    logger.critical('---------------exploding assembly---------------')
    # setup recursion variables
    explode_bom.df_final = pd.DataFrame()
    explode_bom.level = 0
    explode_bom.sort_path = ''
    explode_bom.assy_qty = 0
    explode_bom.assy_qp = 1

    # perform bom recursion
    bom_explosion_df = explode_bom(top_level=top_level_input, make_qty=int(qty_input),
                                   file_path=file_path, ignore_epicor=ignore_epicor)
    bom_explosion_df = bom_explosion_df.rename(columns={'PART NUMBER': 'Part',
                                                        'QTY': 'Comp Q/P'})
    if not bom_explosion_df.empty:

        print('---------------getting supply status---------------')
        logger.critical('---------------getting supply status---------------')
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
        logger.critical('---------------getting floor locations---------------')

        import bin_locations

        # Get floor locations for components
        bin_df = bin_locations.get_bins(bom_explosion_df)

        # Add bin data to df
        bom_explosion_df = bom_explosion_df.merge(bin_df[['Part', 'Main Bins', 'Floor Bins']],
                                                  on='Part', how='left')

        # remove unneeded columns based on arguments passed
        if explode_bom.passed_args['ignore_epicor']:
            bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM'], axis=1)
        else:
            bom_explosion_df = bom_explosion_df.drop(['FUNCTION', 'PhantomBOM', 'sub'], axis=1)

        print('---------------Formatting and writing to excel---------------')
        logger.critical('---------------Formatting and writing to excel---------------')

        bom_explosion_df['Top Level'] = bom_explosion_df['Sort Path'][0][1:]
        bom_explosion_df['BOM Details [BOM Path | Parent Q/P | Comp Q/P]'] = \
            bom_explosion_df[['Sort Path', 'Assembly Q/P', 'Comp Q/P']].astype(str) \
            .apply(lambda x: '[' + '-'.join(x) + ']' + '\n' if x.all != '' else set(x), axis=1)

        bom_explosion_df = bom_explosion_df.rename(columns={'PartDescription': 'Description','TypeCode': 'Type'})
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
        writer = pd.ExcelWriter(r'\\vfile\MPPublic\pdf_bom_explosion'
                                + '\\' + re.escape(assy_part) + f'_{file_date}.xlsx',
                                engine='openpyxl')
        book = writer.book
        pivtbl.to_excel(writer, sheet_name='bom_explosion', index=False)

        # add hyperlink for print
        sheet = book['bom_explosion']
        part_list = pivtbl['Part'].tolist()
        file_path = explode_bom.passed_args['file_path']
        for i, row in enumerate(part_list):
            # sets hyperlink value
            sheet.cell(row=i + 2, column=4).value = \
                f'=HYPERLINK("{file_path}{part_list[i]}.pdf","print")'

            # sets hyperlink format
            sheet.cell(row=i + 2, column=4).font = Font(underline='single', color='0563C1')

        sheet.freeze_panes = 'A2'
        writer.save()

        # open file once process completed
        excel_app = Dispatch("Excel.Application")
        excel_app.Visible = True  # otherwise excel is hidden
        excel_wb = excel_app.Workbooks.Open(r'\\vfile\MPPublic\pdf_bom_explosion'
                                            + '\\' + re.escape(assy_part) + f'_{file_date}.xlsx')

        print('Done!')
        logger.critical('---------------Done! - File Location Below---------------')
        logger.critical(r'\\vfile\MPPublic\pdf_bom_explosion'
                        + '\\' + re.escape(assy_part) + f'_{file_date}.xlsx')

        if email_supp:

            outlook = win32com.client.DispatchEx('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = supp_adr
            mail.CC = 'jboyette@leeboy.com'
            mail.Subject = f'Please quote {top_level_input}'
            mail.Body = ''
            mail.HTMLBody = f"""
                    <body>
                        <p>attched component prints</p>
                    </body>"""

            for part in part_list:
                try:
                    base_url = r'\\vimage\latest' + '\\'
                    part_num = f'{part}.pdf'
                    attachment = base_url + part_num
                    mail.Attachments.Add(attachment)
                except:
                    try:
                        base_url = r'\\vimage\latest' + '\\'
                        part_num = f'{part}.dwg'
                        attachment = base_url + part_num
                        mail.Attachments.Add(attachment)
                    except:
                        pass

            mail.Send()


