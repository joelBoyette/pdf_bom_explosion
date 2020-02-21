
import pdf_bom

import pandas as pd
import warnings
warnings.filterwarnings("ignore")


epicor_data = \
            pd.read_excel(
                r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\test_epicor_data.xlsx',
                sheet_name='data')

subassy_df = pd.read_excel(r'\\Vfile\d\MaterialPlanning\MPPublic\manufactured_part_routings.xlsm')

subassy_df = subassy_df.loc[(subassy_df['ResourceGrpID'] != 'FAB')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'ELEC')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'HOSE')
                            ]
subassy_df = subassy_df['MtlPartNum'].unique()


def explode_bom(top_level, make_qty, file_path=r'\\vimage\latest' + '\\', ignore_epicor=False):

    explode_bom.passed_args = locals()

    explode_bom.part = top_level
    explode_bom.assy_qty = make_qty
    explode_bom.sort_path += '\\' + explode_bom.part

    print(f'traverse bom part {explode_bom.part} '
          f'level {explode_bom.level} path {explode_bom.sort_path} '
          f'assy qty {explode_bom.assy_qty}')

    # Get bom from pdf and merge epicor data
    df = pdf_bom.read_pdf_bom(explode_bom.part, file_path)

    if df.empty:
        pass
    else:

        # use epicor flags to determine which parts to explode
        if not ignore_epicor:

            df = df.merge(epicor_data, on='PART NUMBER', how='left')
            df.loc[df['PART NUMBER'].isin(subassy_df), 'sub'] = 'yes'
            df.loc[~df['PART NUMBER'].isin(subassy_df), 'sub'] = 'no'

            # put non phantoms in df
            df_no_explode = df.loc[(~df['Phantom BOM']) & (df['sub'] == 'no')]
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

            # capture the sort path before recursion.  used to set the sort path on phantoms
            sort_path_reset = explode_bom.sort_path

            # check for phantoms and traverse boms
            df_explode = df.loc[(df['Phantom BOM']) | (df['sub'] == 'yes')][['PART NUMBER', 'QTY']]

        else:

            # do not care about epicor flags, keep exploding as long as prints have BOM's
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

            # capture the sort path before recursion.  used to set the sort path on phantoms
            sort_path_reset = explode_bom.sort_path

            # check for phantoms and traverse boms
            df_explode = df

        if not df_explode.empty:

            for idx, df_row in df_explode.iterrows():

                explode_bom.level += 1

                explode_part = df_explode.loc[idx, 'PART NUMBER']
                explode_qty = int(df_explode.loc[idx, 'QTY'])
                explode_extd_qty = explode_bom.assy_qty * explode_qty
                explode_bom.assy_qp = explode_qty

                # recursion through lower levels
                explode_bom(explode_part, explode_extd_qty, file_path, ignore_epicor)

                # reset bom level and sort path for next traversal
                explode_bom.level -= 1
                explode_bom.sort_path = sort_path_reset
                explode_bom.assy_qty = make_qty

    return explode_bom.df_final


# setup recursion variables
explode_bom.df_final = pd.DataFrame()
explode_bom.level = 0
explode_bom.sort_path = ''
explode_bom.assy_qty = 0
explode_bom.assy_qp = 1
