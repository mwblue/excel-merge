import pandas as pd
import numpy as np
import os
from pdb import set_trace as bp

datapath = './data/'
outpath = './output'


def proc_quarter(qtr_sheet, fin_code_original, largest_asset_code, total_asset_code):
    # bp()
    qtr_tbl = qtr_sheet
    qtr_tbl = qtr_tbl.T
    qtr_tbl.columns = qtr_tbl.loc[0]
    qtr_tbl.drop(0, inplace = True)
    qtr_tbl.reset_index(drop = True, inplace = True)
    
    # covert ASSET SIZE CODE to string
    qtr_tbl['ASSET SIZE CODE'] = qtr_tbl['ASSET SIZE CODE'].astype('str')

    # only keep the columns needed
    #bp()
    fin_code_to_keep = list(set(fin_code_original) & set(qtr_tbl.columns))
    col_to_keep = ['QUARTER & YEAR', 'INDUSTRY CODE', 'ASSET SIZE CODE'] + fin_code_to_keep
    qtr_tbl_0 = qtr_tbl[col_to_keep]  
    qtr_tbl_largest = qtr_tbl_0[qtr_tbl_0['ASSET SIZE CODE']==largest_asset_code]
    qtr_tbl_largest.drop('ASSET SIZE CODE', axis=1, inplace=True)
    qtr_tbl_all  =qtr_tbl_0[qtr_tbl_0['ASSET SIZE CODE']==total_asset_code]
    qtr_tbl_all.drop("ASSET SIZE CODE", axis=1, inplace=True)

    # rename the financial columns for total debt
    new_col_dict = {}
    for col in qtr_tbl_all.columns:
        if col in fin_code_to_keep:
            new_col_dict[col] = col+'-SIZE-ALL'
        else:
            new_col_dict[col] = col
    qtr_tbl_all = qtr_tbl_all.rename(columns = new_col_dict)

    # combine to make new summary table
    # dg()
    qtr_tbl_largest = qtr_tbl_largest.set_index('INDUSTRY CODE')
    qtr_tbl_all = qtr_tbl_all.set_index('INDUSTRY CODE')
    qtr_tbl_new = qtr_tbl_largest.combine_first(qtr_tbl_all)
    qtr_tbl_new = qtr_tbl_new.reset_index()

    quarter = qtr_tbl_new['QUARTER & YEAR']
    qtr_tbl_new.drop(labels=['QUARTER & YEAR'], axis=1,inplace = True)
    qtr_tbl_new.insert(0, 'QUARTER & YEAR', quarter)

   ## qtr_tbl_new = qtr_tbl_new.rename(columns={'STBANK':'STBANK-SIZE-LARGEST'})
   ## qtr_tbl_new = qtr_tbl_new[['QUARTER & YEAR', 'INDUSTRY CODE','STBANK-SIZE-ALL', 'STBANK-SIZE-LARGEST']]
    
    qtr_tbl_new.reset_index(drop=True, inplace=True)
    qtr_tbl_new.fillna("", inplace=True)
    return qtr_tbl_new

def proc_one_xlsfile(in_name, out_fname, fin_code, year):
    #bp()
    df = pd.read_excel(in_name, header = None, sheet_name = None )
    df_key = df['KEY']
    key_names = [x for x in df.keys()]
    key_names = key_names[1:]
    largest_asset_code, total_asset_code = get_code(df_key)
    if year == '1947':
        qtr_list = ['1947Q2', '1947Q3','1947Q4']
    else:
        qtr_list = key_names
    summary = []
    for q in qtr_list:
        qtr = df[q]
        q_sum = proc_quarter(qtr, fin_code, largest_asset_code, total_asset_code)
        summary.append(q_sum) 
    y_sum = pd.concat(summary, axis=0)
    y_sum.reset_index(inplace=True)
    y_sum.drop("index", axis = 1, inplace = True)
    y_sum.to_excel(out_fname, index = False)

def get_code(df_key):
    # get the largest asset code and total asset code 
    start_idx = df_key.index[df_key[0]=='Asset Size Code'].tolist()
    end_idx = df_key.index[df_key[0]=='Financial Data Item Code Definitions'].tolist()
    codes = df_key.iloc[start_idx[0]+1: end_idx[0]-1,0].tolist()
    int_codes = []
    for v in codes:
        if not np.isnan(np.float(v)):
            int_codes.append(int(v))
    int_codes.sort()
    total_asset_code = format(int_codes[-1], '02d')
    largest_asset_code = format(int_codes[-2], '02d')	
    return largest_asset_code, total_asset_code


""" Test """
# bp()
files_list = os.listdir(datapath)
for fname in files_list:
    year = fname.split('.')[0][-4:]
    fname = os.path.join(datapath, fname)
    out_fname = os.path.join(outpath, year+"_summary.xls")
    fin_code = ['STBANK', 'COMPAPER', 'STDEBTOTH', 'INSTBANKS', 'INSTBONDS', 'INSTOTHER', 'LTBNKDEBT', 'LTBNDDEBT', 'LTOTHDEBT']

    proc_one_xlsfile(fname, out_fname, fin_code, year)
    #"""
    results = pd.read_excel(out_fname, header=None, index_col=0)
    print("Results:")
    print(results)
    #"""
