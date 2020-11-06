import logging
import camelot
import tabula
import re
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings("ignore")

logger = logging.getLogger(__name__)


# #################### TABULA PY  ############################
# https://tabula-py.readthedocs.io/en/latest/getting_started.html#example
def clean_pdf_tabulapy(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = tabula.read_pdf(full_path, stream=True, pages=1)
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = tabula.read_pdf(default_path, stream=True, pages=1)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():
        #     print('Column Name : ', column_name)
        #     print('Column Contents : ', column_data.values)
        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 3]
    df_final = df_final.dropna(how='all')

    part_list = []
    final_list = []
    if clean_type == 'PART NUMBER':

        headers = list(df_final.iloc[0])  # put headers in a list
        headers = [str(header).replace('QTY.', 'QTY') for header in headers]

        df_final.columns = headers  # #makes columns match the headers list
        df_final = df_final.iloc[1:]  # remove rows before headers

        df_final = df_final.fillna('-')

        df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
            '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|[0-9]', regex=True)]

        for i, row in df_final.iterrows():

            part = df_final.loc[i, 'PART NUMBER']
            qty = df_final.loc[i, 'QTY']

            search_result = re.search('(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]', part)

            if search_result is not None:
                split_part = part.split(' ')

                split_qty = qty.split(' ')

                part_dict = {'PART NUMBER': split_part[0], 'QTY': split_qty[0]}
                part_list.append(part_dict)
                part_dict = {'PART NUMBER': split_part[1], 'QTY': split_qty[1]}
                part_list.append(part_dict)

                df_final = df_final.loc[df_final['PART NUMBER'] != part]

        df_temp = pd.DataFrame(part_list)
        df_final = pd.concat([df_final, df_temp], sort=False)

        df_final = df_final.loc[(~df_final['PART NUMBER'].str.contains('[0-9] [0-9]|\d[.]'))
                                & (df_final['QTY'] != '-')]

        df_final = df_final[['PART NUMBER', 'QTY']]
        df_final['QTY'] = df_final['QTY'].str.extract(r'(\d+)', expand=False).str.strip()
        df_final = df_final.dropna(how='any')
        df_final['clean_type'] = clean_type

    elif clean_type == 'PART NUMBER QTY':

        headers = list(df_final.iloc[0])
        headers = [str(header).replace('QTY.', 'QTY') for header in headers]
        df_final.columns = headers
        df_final = df_final.dropna(how='any')
        df_final = df_final.rename(columns={'PART NUMBER QTY': 'PART NUMBER'})
        df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
            '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|[0-9]', regex=True)]

        regex_expr = re.compile('(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]')

        for part_row in df_final['PART NUMBER']:

            regex_results = re.findall(regex_expr, part_row)

            if len(regex_results) > 1:
                for result in regex_results:
                    split_result = result.split(' ')
                    part_list += split_result
                reorder = [0, 2, 1, 3]
                part_list = [part_list[i] for i in reorder]

                new_list = [part_list[0] + '-' + part_list[1], part_list[2] + '-' + part_list[3]]
                for item in new_list:
                    part_qty = item.split('-')
                    part_dict = {'PART NUMBER': part_qty[0], 'QTY': part_qty[1]}
                    final_list.append(part_dict)
            else:
                for result in regex_results:
                    split_result = result.split(' ')
                    part_dict = {'PART NUMBER': split_result[0], 'QTY': split_result[1]}
                    final_list.append(part_dict)

        df_final = pd.DataFrame(final_list)
        df_final = df_final[['PART NUMBER', 'QTY']]
        df_final['clean_type'] = clean_type

    elif clean_type == 'NO. PART NUMBER':

        headers = list(df_final.iloc[1])
        headers = [str(header).replace('QTY.', 'QTY') for header in headers]
        df_final.columns = headers
        df_final = df_final.iloc[2:]
        df_final = df_final.dropna(how='any')
        df_final = df_final.rename(columns={'NO. PART NUMBER': 'PART NUMBER'})

        df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
            '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|[0-9]', regex=True)]

        df_final['PART NUMBER'] = df_final['PART NUMBER'].apply(lambda x: x.split(' ')[1])
        df_final = df_final[['PART NUMBER', 'QTY']]
        df_final['clean_type'] = clean_type
    elif clean_type == 'NUMBER':
        headers = list(df_final.iloc[0])  # put headers in a list
        headers = [str(header).replace('QTY.', 'QTY') for header in headers]

        df_final.columns = headers  # #makes columns match the headers list
        df_final = df_final.iloc[1:]  # remove rows before headers

        df_final = df_final.fillna('-')
        df_final = df_final.rename(columns={'NUMBER': 'PART NUMBER'})
        df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
            '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|[0-9]', regex=True)]

        for i, row in df_final.iterrows():

            part = df_final.loc[i, 'PART NUMBER']
            qty = df_final.loc[i, 'QTY']

            search_result = re.search('(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]', part)

            if search_result is not None:
                split_part = part.split(' ')

                split_qty = qty.split(' ')

                part_dict = {'PART NUMBER': split_part[0], 'QTY': split_qty[0]}
                part_list.append(part_dict)
                part_dict = {'PART NUMBER': split_part[1], 'QTY': split_qty[1]}
                part_list.append(part_dict)

                df_final = df_final.loc[df_final['PART NUMBER'] != part]

        df_temp = pd.DataFrame(part_list)
        df_final = pd.concat([df_final, df_temp], sort=False)

        df_final = df_final.loc[(~df_final['PART NUMBER'].str.contains('[0-9] [0-9]|\d[.]'))
                                & (df_final['QTY'] != '-')]

        df_final = df_final[['PART NUMBER', 'QTY']]
        df_final['QTY'] = df_final['QTY'].str.extract(r'(\d+)', expand=False).str.strip()
        df_final = df_final.dropna(how='any')
        df_final['FUNCTION'] = 'tabulapy'

        return df_final

    df_final = pd.DataFrame(columns=['PART NUMBER', 'QTY'])
    df_final['clean_type'] = full_path

    return df_final


# ################ CAMELOT ####################
# https://camelot-py.readthedocs.io/en/master/user/advanced.html
def clean_pdf_camelot_scale15(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        # pdf_df = camelot.read_pdf(full_path, line_scale=15, shift_text=[''])[0].df
        pdf_df = camelot.read_pdf(full_path, line_scale=15, shift_text=[''])[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=15)[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)
    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]

    if df_final.iloc[4, 0] != '' or df_final.iloc[5, 0] != '':
        df_final = df_final.dropna(how='all')
        df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)
    else:
        raise Exception

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'}).reset_index(drop=True)

    if df_final['PART NUMBER'][0] == '':
        raise Exception

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains('(^[.])|([\/])', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]

    df_final['FUNCTION'] = 'camelot_scale15'

    return df_final


def clean_pdf_camelot_scale10(part, file_path):

    full_path = file_path + f'{part}.pdf'
    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=10, shift_text=[''])[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=10, shift_text=[''])[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)
    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]

    if df_final.iloc[4, 0] != '':

        df_final = df_final.dropna(how='all')
        df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)
    else:
        raise Exception

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list

    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'}).reset_index(drop=True)

    if df_final['PART NUMBER'][0] == '':
        raise Exception

    df_final = df_final.loc[(
            df_final['PART NUMBER'].str.contains('(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True))
            & (df_final['QTY'].str.isdigit())]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]

    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale10'

    return df_final


def clean_pdf_camelot_scale5(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=5, shift_text=[''])[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=5, shift_text=[''])[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']
    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]

    if df_final.iloc[4, 0] != '':

        df_final = df_final.dropna(how='all')
        df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)
    else:
        raise Exception

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'}).reset_index(drop=True)

    if df_final['PART NUMBER'][0] == '':
        raise Exception

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]

    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale5'

    return df_final


def clean_pdf_camelot_stream(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, flavor='stream')[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, flavor='stream')[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]
    df_final = df_final.dropna(how='all')
    df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'})

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_stream'

    return df_final


def clean_pdf_camelot_scale20(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=20, shift_text=[''])[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=20, shift_text=[''])[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            if 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']
    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]

    if df_final.iloc[4, 0] != '':

        df_final = df_final.dropna(how='all')
        df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)
    else:
        raise Exception

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'}).reset_index(drop=True)

    if df_final['PART NUMBER'][0] == '':
        raise Exception

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale20'

    return df_final


def clean_pdf_camelot_scale15_split_text(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=15, shift_text=[''], split_text=True)[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=15, shift_text=[''], split_text=True)[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]

    if df_final.iloc[4, 0] != '':

        df_final = df_final.dropna(how='all')
        df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)
    else:
        raise Exception

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'}).reset_index(drop=True)

    if df_final['PART NUMBER'][0] == '':
        raise Exception

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale15_splittext'

    return df_final


def clean_pdf_camelot_scale10_split_text(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=10, shift_text=[''], split_text=True)[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=10, shift_text=[''], split_text=True)[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            if 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]
    df_final = df_final.dropna(how='all')
    df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'})

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale10_splittext'

    return df_final


def clean_pdf_camelot_scale5_split_text(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=5, shift_text=[''], split_text=True)[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=5, shift_text=[''], split_text=True)[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']
    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]
    df_final = df_final.dropna(how='all')
    df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'})

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']
    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale5_splittext'

    return df_final


def clean_pdf_camelot_scale20_split_text(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=20, shift_text=[''], split_text=True)[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=20, shift_text=[''], split_text=True)[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)
    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            if 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]
    df_final = df_final.dropna(how='all')
    df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'})

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']

    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale20_splittext'

    return df_final


def clean_pdf_camelot_scale3(part, file_path):

    full_path = file_path + f'{part}.pdf'

    try:
        pdf_df = camelot.read_pdf(full_path, line_scale=3, shift_text=[''])[0].df
    except:
        # if user supplied path doesnt work
        default_path = r'\\vimage\latest' + '\\' + f'{part}.pdf'
        pdf_df = camelot.read_pdf(default_path, line_scale=3, shift_text=[''])[0].df

    pdf_df = pdf_df.replace('\\n', '', regex=True)

    find_column_dict = {}
    for (column_name, column_data) in pdf_df.iteritems():

        for x in column_data.values:
            if 'ASTM' in str(x):
                raise Exception
            elif 'PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER'}
            elif 'PART NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'PART NUMBER QTY'}
            elif 'NO. PART NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NO. PART NUMBER'}
            elif 'NUMBER' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER'}
            elif 'NUMBER QTY' == str(x):
                find_column_dict = {'column': column_name, 'clean type': 'NUMBER QTY'}

    clean_type = find_column_dict['clean type']

    col_index = pdf_df.columns.get_loc(find_column_dict['column'])
    df_final = pdf_df.iloc[:, col_index:col_index + 6]
    df_final = df_final.dropna(how='all')
    df_final = df_final[df_final.iloc[:, 0] != ''].reset_index(drop=True)

    if clean_type == 'PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER'].index[0]
    elif clean_type == 'PART NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'PART NUMBER QTY'].index[0]
    elif clean_type == 'NO. PART NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NO. PART NUMBER'].index[0]
    elif clean_type == 'NUMBER':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER'].index[0]
    elif clean_type == 'NUMBER QTY':
        header_index = df_final[df_final.iloc[:, 0] == 'NUMBER QTY'].index[0]

    headers = list(df_final.iloc[header_index])  # put headers in a list
    headers = [str(header).replace('QTY.', 'QTY') for header in headers]
    df_final.columns = headers  # #makes columns match the headers list
    df_final = df_final.iloc[header_index + 1:]  # remove rows that contained the headers
    df_final = df_final.rename(columns={df_final.columns[0]: 'PART NUMBER'})

    df_final = df_final.loc[df_final['PART NUMBER'].str.contains(
        '(\d+[A-Z] )|([A-Z]+\d)[\dA-Z]*|[0-9] [0-9]|\d\d+|\d[.]-', regex=True)]

    df_final = df_final.loc[~df_final['PART NUMBER'].str.contains(' ', regex=True)]
    if 'QTY' not in df_final.columns:
        df_final = df_final.replace(' ', np.nan)
        df_final = df_final.dropna(axis=1, how='all')
        df_final = df_final.iloc[:, 0:2]
        df_final.columns = ['PART NUMBER', 'QTY']

    df_final = df_final[['PART NUMBER', 'QTY']]
    df_final['FUNCTION'] = 'camelot_scale3'

    return df_final


def read_pdf_bom(part, file_path=r'\\vimage\latest' + '\\'):

    pdf_bom_all_df = pd.DataFrame()

    for pdf_func in [clean_pdf_tabulapy, clean_pdf_camelot_scale15, clean_pdf_camelot_scale10,
                     clean_pdf_camelot_scale5, clean_pdf_camelot_stream, clean_pdf_camelot_scale20,
                     clean_pdf_camelot_scale15_split_text, clean_pdf_camelot_scale10_split_text,
                     clean_pdf_camelot_scale5_split_text, clean_pdf_camelot_scale20_split_text,
                     clean_pdf_camelot_scale3
                     ]:
        try:
            pdf_bom_df = pdf_func(part, file_path)
            any_blanks = pdf_bom_df.loc[:, (pdf_bom_df == '').all()].count().empty
            alpha_qty = pdf_bom_df.loc[pdf_bom_df['QTY'].str.contains('[A-Z]', regex=True)].count()['QTY']

            if pdf_bom_df.empty:
                raise Exception
            elif not any_blanks:
                raise Exception
            elif alpha_qty > 0:
                raise Exception

        except Exception as e:
            # print(f' no methods worked for {part}')
            pdf_bom_df = pd.DataFrame(
                columns=['PART NUMBER', 'QTY', 'FUNCTION'])
        else:
            break

    pdf_bom_df['Assembly'] = part

    # appends current ecn df to the overall df
    pdf_bom_all_df = pdf_bom_all_df.append(pdf_bom_df, sort=False)

    pdf_bom_all_df = pdf_bom_all_df.loc[~pdf_bom_all_df['PART NUMBER'].str.contains(
        '( +)|(^\d$)|(^[.])', regex=True)]

    pdf_bom_all_df = pdf_bom_all_df.loc[~pdf_bom_all_df['PART NUMBER'].str.contains(
        ' ', regex=True)]

    pdf_bom_all_df = pdf_bom_all_df.loc[~pdf_bom_all_df['QTY'].str.contains(
        r'^\s*$', regex=True)]

    return pdf_bom_all_df

