import pandas as pd


def get_bins(part_df):

    part_df['Part'] = part_df['Part'].astype(str)

    bins_df = pd.read_excel(r'\\vfile\MPPublic\Projects\NPD\NPD_data.xlsm', sheet_name='bin_contents')
    bins_df['Bin'] = bins_df['Bin'].astype(str)
    bins_df['Bin'] = bins_df['Bin'].str.upper()
    bins_df = bins_df.rename(columns={'OH': 'Bin OH'})

    main_df = bins_df.loc[(bins_df['WH'] == 'MAIN')
                          | (bins_df['WH'] == 'RDZ')
                          | (bins_df['WH'] == 'STAR')]

    floor_df = bins_df.loc[(bins_df['WH'] != 'MAIN')
                           & (bins_df['WH'] != 'SHIP')
                           & (bins_df['WH'] != 'RDZ')
                           & (bins_df['WH'] != 'STAR')]

    # combine main bins detail
    if not main_df.empty:

        part_df = part_df.merge(main_df[['Part', 'WH', 'Bin', 'Bin OH']], on='Part', how='left')
        part_df = part_df.fillna('')

        part_df['Bin OH'] = part_df['Bin OH'].astype(str)
        part_df['Main Bins'] = part_df[['Bin', 'Bin OH']] \
                               .apply(lambda x: ' ['.join(x) + '] ' if x.all != '' else set(x), axis=1)

        part_df['Main Bins'] = part_df['Main Bins'].str.replace('\[]', '')
        part_df = part_df.drop(['WH', 'Bin', 'Bin OH'], axis=1)

    # combine floor bins detail
    if not floor_df.empty:

        part_df = part_df.merge(floor_df[['Part', 'WH', 'Bin', 'Bin OH']], on='Part', how='left')
        part_df = part_df.fillna('')

        part_df['Bin OH'] = part_df['Bin OH'].astype(str)
        part_df['Floor Bins'] = part_df[['Bin', 'Bin OH']] \
            .apply(lambda x: ' ['.join(x) + '] ' if x.all != '' else set(x), axis=1)

        part_df['Floor Bins'] = part_df['Floor Bins'].str.replace('\[]', '')
        part_df = part_df.drop(['WH', 'Bin', 'Bin OH'], axis=1)

    cols = part_df.columns[~part_df.columns.isin(['Part'])].tolist()
    pivtbl = pd.pivot_table(part_df,
                            index=['Part'],
                            values=cols,
                            aggfunc=lambda x: ''.join(list(set(x))),
                            fill_value='').reindex(columns=cols).reset_index()

    return pivtbl

