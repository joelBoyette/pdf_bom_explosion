import pandas as pd

subassy_df = pd.read_excel(r'\\Vfile\d\MaterialPlanning\MPPublic\manufactured_part_routings.xlsm')

subassy_df = subassy_df.loc[(subassy_df['ResourceGrpID'] != 'FAB')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'ELEC')
                            & (subassy_df['ResourceGrpID'].str[0:4] != 'HOSE')
                            ]
subassy_df = subassy_df['MtlPartNum'].unique()

bom = pd.read_excel(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\output.xlsx',
                    sheet_name='bom_explosion')

bom.loc[bom['PART NUMBER'].isin(subassy_df), 'sub'] = 'yes'
bom.loc[~bom['PART NUMBER'].isin(subassy_df), 'sub'] = 'no'

df_no_phantom = bom.loc[(~bom['Phantom BOM']) | (bom['sub'] == 'no')]

df_phantoms = bom.loc[(bom['Phantom BOM']) | (bom['sub'] == 'yes')][['PART NUMBER', 'QTY']]

# print(bom.loc[bom['Phantom BOM']])
#
# print(bom.loc[(bom['sub'] == 'yes')])
print(df_phantoms)