# -*- coding: utf-8 -*-
"""
Created on Tue Mar 23 13:08:19 2021

@author: Kianoush
"""

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
import statistics
from statsmodels.stats.outliers_influence import OLSInfluence

############################# Market Model ###################################

path = r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Tickers.xlsx'
Block_Trades_df = pd.read_excel(path)
Block_Trades_df['DateMiladi'] = pd.to_datetime(Block_Trades_df['DateMiladi'], format='%Y%m%d')
Block_Trades_df['DateBegin'] = pd.to_datetime(Block_Trades_df['DateBegin'], format='%Y%m%d')

market_df = pd.read_excel(r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\IndexData.xls') 
market_df['market_return'] = np.log(market_df['<CLOSE>']).diff(20) 
market_df['<DTYYYYMMDD>'] = pd.to_datetime(market_df['<DTYYYYMMDD>'], format='%Y%m%d')


dict = {} 
for i in Block_Trades_df.index:
    dict[i] = np.nan
 
for i in Block_Trades_df.index: 
    print(i)
    ticker = Block_Trades_df['Ticker'].iloc[i] 
    IDX = Block_Trades_df["IDX"].iloc[i]
    date = Block_Trades_df['DateMiladi'].iloc[i] 
    
    path = r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Price'
    #f_path = path + '\\' + ticker + '.csv' 
    # f_path_1 = path + '\\' + IDX + '.csv' 
    f_path_2 = path + '\\' + str(IDX) + '.xlsx' 
    try:
        temp_df = pd.read_excel(f_path_2)
    except:
        continue
#    try:
#        temp_df = pd.read_csv(f_path, encoding='ISO-8859–1')
#    except:
#        temp_df = pd.read_excel(r'C:\Users\PC-1\Desktop\All Tickers' + '\\' + ticker + '.xls')
    
    temp_df['<DTYYYYMMDD>'] = pd.to_datetime(temp_df['<DTYYYYMMDD>'], format='%Y%m%d')
    dict[i] = temp_df[['<DTYYYYMMDD>', '<CLOSE>']]
    dict[i]['return'] = np.log(dict[i]['<CLOSE>']).diff(20) 
    list_of_columns = ['Market Model Alpha','Market Model Beta', 'std']  
    list_of_values = [np.nan]*len(list_of_columns)
    dict[i][list_of_columns] = pd.DataFrame([list_of_values], columns = list_of_columns,
                                             index = dict[i].index)

    stop_date = date 

    reg_window = 1205 
    stop_index_i = int(np.where(dict[i]['<DTYYYYMMDD>'] == stop_date)[0]) 
    stop_index_m = int(np.where(market_df['<DTYYYYMMDD>'] == stop_date)[0]) 
    if stop_index_i > reg_window:
        start_index_i = stop_index_i - reg_window 
        start_index_m = stop_index_m - reg_window 
        if start_index_m <= 0:
            print("Error1:", ticker)
            dict[i]=np.nan
            continue
    else:
        if stop_index_i >= reg_window/5:
            start_index_i = 1 
            start_index_m = stop_index_m - (stop_index_i - start_index_i)
            if start_index_m <= 0:
                print("Error2:", ticker)
                dict[i]=np.nan
                continue
        else:
            print("Error3:", ticker)
            dict[i]=np.nan
            continue

    y = np.array(dict[i]['return'].iloc[start_index_i : stop_index_i])
    x = np.array(market_df['market_return'].iloc[start_index_m : stop_index_m])
    d = { "x": pd.Series(x), "y": pd.Series(y)}
    df = pd.DataFrame(d)

    # Market model
    mod = smf.ols('y ~ x', data=df)
    reg = mod.fit() 
    std = OLSInfluence(reg).resid_std.mean()
    dict[i]['Market Model Beta'].iloc[stop_index_i] = reg.params[1] 
    dict[i]['Market Model Alpha'].iloc[stop_index_i] = reg.params[0] 
    y_prediction = reg.params[0] + reg.params[1]*market_df['market_return'].iloc[stop_index_m]
    y_actual = dict[i]['return'].iloc[stop_index_i]
    dict[i]['std'].iloc[stop_index_i] = std
                    
    dict[i] = dict[i][dict[i]['Market Model Beta'].notna()]
    dict[i] = dict[i].reset_index(drop=True)
    
#
new_dict = {} 
for i in dict.keys():
    if isinstance(dict[i], pd.DataFrame):
        new_dict[i] = dict[i]
dict = new_dict

#
output = pd.DataFrame()
for key in dict.keys():
    output = output.append(dict[key][['Market Model Alpha','Market Model Beta', 'error']], ignore_index=True)
output["IDX"] = dict.keys()
output.set_index("IDX", drop=True, inplace=True)

############################# End of Code ###################################

market_sigma2 = statistics.variance(market_df["market_return"].iloc[1265:2471])




