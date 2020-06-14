import pandas as pd
import numpy as np
import os
from pdb import set_trace as bp

datapath = './files/'
outpath = './output'

# bp()
files_list = os.listdir(datapath)
all_data=pd.DataFrame()
for fname in files_list:
    fname = os.path.join(datapath, fname)
    bp()
    df=pd.read_excel(fname)
    all_data=all_data.append(df,ignore_index=True)
""" put QUATER & YEAR first """
cols=all_data.columns.tolist()
cols.remove('QUARTER & YEAR')
new_cols=['QUARTER & YEAR']+cols
all_data=all_data[new_cols]

writer=pd.ExcelWriter('output.xlsx')
all_data.to_excel(writer,'sheet1')
writer.save()
