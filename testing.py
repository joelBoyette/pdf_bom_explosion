import pandas as pd

part = 'beef'

pdf_bom_df = pd.DataFrame([[part, 0, 'DIDNT WORK']],
                        columns=['PART NUMBER', 'QTY', 'FUNCTION'])


print(pdf_bom_df)