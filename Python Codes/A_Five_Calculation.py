# -*- coding: utf-8 -*-
"""
Created on Sat Aug 14 22:16:33 2021

@author: Kianoush
"""

import numpy as np
import pandas as pd

path = r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Tickers.xlsx'
Block_Trades_df = pd.read_excel(path)
Block_Trades_df['DateMiladi'] = pd.to_datetime(Block_Trades_df['DateMiladi'], format='%Y%m%d')
Block_Trades_df['DateBegin'] = pd.to_datetime(Block_Trades_df['DateBegin'], format='%Y%m%d')

def Average(lst):
    return sum(lst) / len(lst)

Block_Trades_df["AFive"]=np.nan
error_list = []
for i in Block_Trades_df.index:
    print(i)
    avg_list = []
    path = r"C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Large Shareholders\Ownership"
    df = pd.read_excel(path + "\\" + "{}.xlsx".format(Block_Trades_df["IDX"].iloc[i]))
    if len(df) == 0:
        error_list.append(i)
        continue
    for j in df.index:
        sh_shares = list(df.iloc[j,1:].dropna().values)
        sh_shares.sort(reverse = True)
        if len(sh_shares) >= 5:
            avg_list.append(Average(sh_shares[0:5]))
        elif len(sh_shares) > 0:
            avg_list.append(Average(sh_shares[:]))
        else:
            avg_list.append(np.nan)
    avg_list_withot_nan = [x for x in avg_list if not np.isnan(x)]
    Block_Trades_df["AFive"].iloc[i] = Average(avg_list_withot_nan)
    
AFive_df = Block_Trades_df.loc[:, ["IDX", "Ticker", "AFive"]]
AFive_df.to_excel(r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\AFive_data.xlsx', index=False)
