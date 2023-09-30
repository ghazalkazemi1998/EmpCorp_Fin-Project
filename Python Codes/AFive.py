# -*- coding: utf-8 -*-
"""
Created on Mon Nov  9 15:41:11 2020

@author: Kianoush
"""

import numpy as np
import pandas as pd
import dateConvertor as dc
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException
import matplotlib.pyplot as plt
import time
from pandas import Series
#################################################################################################
#################################################################################################
                                            ### START ###
                                            
market_df = pd.read_excel(r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Large Shareholders\IndexData.xls')
all_dates_list = list(market_df['<DTYYYYMMDD>']) 
Tickers = pd.read_excel(r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Large Shareholders\Tickers.xlsx') 
for i in Tickers.index:
    Tickers['DateMiladi'].iloc[i] = str(Tickers['DateMiladi'][i]).split('.')[0]
    Tickers['DateBegin'].iloc[i] = str(Tickers['DateBegin'][i]).split('.')[0]
    
not_nan_idx = [] 
for idx in Tickers.index:
        not_nan_idx.append(idx)

main_dict = {} 
for idx in not_nan_idx:
    id_company = Tickers['Company ID'].iloc[idx][1:-1] 
    trade_date = Tickers['DateMiladi'].iloc[idx] 
    begin_date = Tickers['DateBegin'].iloc[idx]
    
    date_list_idx = [idx for idx in range(len(all_dates_list)) 
                        if (all_dates_list[idx] <= int(trade_date) and all_dates_list[idx] >= int(begin_date))] 
    period = 20 
    date_list = [str(all_dates_list[i]) for i in date_list_idx if i % period == 0] 
    dict={}
    id_dict={} 
    
    is_close = 1 
    for date in date_list:
        print(idx)
        print(date)
        time.sleep(2)
        if is_close==1:
            driver = webdriver.Chrome()
            driver.set_page_load_timeout(15) 
            is_close=0
        try:    
            driver.get("http://cdn.tsetmc.com/Loader.aspx?ParTree=15131P&i=" + id_company +"&d=" + date)
        except:
            print("TimeOut Error")
            df = pd.DataFrame() 
            dict[date] = df.copy()
            id_dict[date] = df.copy()
            driver.quit() 
            is_close = 1
            continue
        
        try:
            driver.find_elements_by_id('shtt')
        except:
            pass
    
        try:
            driver.find_elements_by_id('shtt') 
        except TimeoutException:
            print("TimeOut Error")
            df = pd.DataFrame() 
            dict[date] = df.copy()
            id_dict[date] = df.copy()
            driver.quit() 
            is_close = 1
            continue 
            
        try: 
            shareholders_table = driver.find_elements_by_id('shtt')[0].text.split('\n')
            name_list = [] 
            share_list = [] 
            shareholder_id = [] 
            
            buttons = driver.find_elements_by_xpath('//div/form/div/div/div/div/div/div/table/tbody[@id="shtt"]/tr[@class="sh"]') 
            onclick_text = [] 
            for b in buttons:   
                onclick_text.append(b.get_attribute('onclick')) 
            for i in range(len(onclick_text)):    
                shareholder_id.append(onclick_text[i].split(",")[0].split("(")[1].split("'")[-1])
            
            for i in range(len(shareholders_table)):
                name_list.append(' '.join(shareholders_table[i].split(' ')[:-3])) 
                share_list.append(float(shareholders_table[i].split(' ')[-1]))

            df = pd.DataFrame(np.nan, columns = shareholder_id, index = range(1))
            df.loc[0] = share_list 
            dict[date] = df.copy()
            id_dict[date] = pd.DataFrame(list(zip(name_list, shareholder_id)), columns = ['name', 'id'],
                                            index = range(len(shareholder_id))).copy() 
        except: 
            df = pd.DataFrame() 
            dict[date] = df.copy()
            id_dict[date] = df.copy()
            print("No Info. Exist")        
            pass
    
    
    temp_shareholders_list = []    
    for key in dict:
        temp_shareholders_list = temp_shareholders_list + list(dict[key].columns)
          
    shareholders_list = [] 
    [shareholders_list.append(x) for x in temp_shareholders_list if x not in shareholders_list] 
    
    main_df = pd.DataFrame(np.nan, columns = ['Date'] + shareholders_list, index = range(len(date_list)))
    for s_h in shareholders_list:
        i = 0
        for key in dict:
            main_df['Date'].iloc[i] = key 
            try:
                main_df[s_h].iloc[i] = dict[key][s_h].iloc[0].copy() 
            except:
                main_df[s_h].iloc[i] = np.nan 
            i+=1
    main_df.to_excel(r'C:\Users\Kianoush\Desktop\Ownership\{}.xlsx'.format(idx), index = False)
    main_dict[trade_date] = main_df.copy()
    
    id_df = pd.DataFrame(columns = ['name','id']) 
    for key in id_dict:
        id_df = id_df.append(id_dict[key])
    id_df = id_df.reset_index(drop=True) 
    id_df = id_df.drop_duplicates(subset='id', keep="first")
    id_df = id_df.reset_index(drop=True)
    
    id_df.to_excel(r'C:\Users\Kianoush\Desktop\Shareholders\{}.xlsx'.format(idx), index = False)
    if is_close==0:
        driver.quit()
        is_close = 1
    
#################################################################################################
#################################################################################################
                                            ### END ###
