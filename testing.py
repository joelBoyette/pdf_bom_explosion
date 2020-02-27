import pandas as pd
import re
import numpy as np

string = r'\1021045\1021338\1016217'

bom_explosion_df = \
    pd.read_excel(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\1021747_02-26-2020.xlsx')

bom_explosion_df['Top Level'] = bom_explosion_df['Sort Path'][0][1:]
bom_explosion_df['BOM Details [BOM Path | Parent Q/P | Comp Q/P]'] = bom_explosion_df[['Sort Path', 'Assembly Q/P',
                                                                                       'Comp Q/P']].astype(str) \
                                                    .apply(lambda x: '[' + '-'.join(x) + ']' + '\n'
                                                            if x.all != '' else set(x), axis=1)

bom_explosion_df = bom_explosion_df.drop(['Assembly', 'Assembly Make Qty', 'Comp Extd Qty',
                                          'Sort Path', 'Level', 'Comp Q/P', 'Assembly Q/P'], axis=1)

bom_explosion_df = bom_explosion_df.rename(columns={'PartDescription': 'Description',
                                                    'TypeCode': 'Type'})

bom_explosion_df = bom_explosion_df[['Part', 'Description', 'Type', 'Dwg Link', 'Total Needed', 'Supply Status',
                                     'Main Bins', 'Floor Bins', 'OH', 'OH_Inspect', 'ONO', 'First Due', 'DMD', 'Buyer',
                                     'Cost', 'Top Level', '# Top Level to Make',
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

print(pivtbl.head(5))

pivtbl.to_excel(r'C:\Users\JBoyette.BRLEE\Documents\Development\test_data\pdf_bom_explosion\test.xlsx')