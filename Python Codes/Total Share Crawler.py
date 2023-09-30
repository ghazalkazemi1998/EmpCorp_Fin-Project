
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
import time

#################################################################################################
#################################################################################################
                                            ### START ###
                                            
market_df = pd.read_excel(r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\IndexData.xls') 
all_dates_list = list(market_df['<DTYYYYMMDD>'])
Tickers = pd.read_excel(r'C:\Users\Kianoush\Desktop\Courses\Empirical Corporate Finance\Project\Data\Tickers.xlsx') 
for i in Tickers.index:
    Tickers['DateMiladi'].iloc[i] = str(Tickers['DateMiladi'][i]).split('.')[0]
    Tickers['DateBegin'].iloc[i] = str(Tickers['DateBegin'][i]).split('.')[0]
    
not_nan_idx = [] 
for idx in Tickers.index:
        not_nan_idx.append(idx)

main_dict = {} 
for idx in not_nan_idx[:]:
    id_company = Tickers['Company ID'].iloc[idx][1:-1] 
    trade_date = Tickers['DateMiladi'].iloc[idx] 
    begin_date = Tickers['DateBegin'].iloc[idx]
    
    date_list_idx = [idx for idx in range(len(all_dates_list)) 
                        if (all_dates_list[idx] <= int(trade_date) and all_dates_list[idx] >= int(begin_date))]
    period = 20 
    date_list = [str(all_dates_list[i]) for i in date_list_idx if i % period == 0]
    dict={} 
    
    is_close = 1 
    for date in date_list:
        print(idx)
        print(date)
        time.sleep(1.5)
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
            driver.quit() 
            is_close = 1
            continue
                
        try:
            driver.find_elements_by_class_name('currencyformat')
        except:
            pass
    
        try:
            driver.find_elements_by_class_name('currencyformat')
        except TimeoutException:
            print("TimeOut Error")
            df = pd.DataFrame()
            dict[date] = df.copy()
            driver.quit() 
            is_close = 1
            continue
            
        try:
            time.sleep(2)
            driver.find_elements_by_class_name('currencyformat')[0]
            total_share_string = ' '.join(driver.find_elements_by_class_name('currencyformat')[0].text.split(" ")[2:])
            total_share_num = float(total_share_string.split(" ")[0])
            total_share_multiplier = 10**6 if total_share_string.split(" ")[1]=='M' else 10**9
            total_shares = total_share_num*total_share_multiplier
            df = pd.DataFrame(np.nan, columns = ['total shares'], index = range(1)) 
            df.loc[0] = total_shares
            dict[date] = df.copy()
        except:
            df = pd.DataFrame() 
            dict[date] = df.copy()
            print("No Info. Exist")        
            pass
    
    
    main_df = pd.DataFrame(np.nan, columns = ['Date', 'total shares'], index = range(len(date_list))) 
    main_df['Date'] = pd.Series(date_list)
    counter = 0
    for key in dict:
        if len(dict[key])>0:
            main_df['total shares'].iloc[counter] = dict[key]['total shares'].iloc[0]
        counter += 1

    main_df.to_excel(r'C:\Users\Kianoush\Desktop\total share\{}.xlsx'.format(idx), index = False)
    main_dict[idx] = main_df.copy()
    
    if is_close==0:
        driver.quit()
        is_close = 1
    
#################################################################################################
#################################################################################################
                                            ### END ###
