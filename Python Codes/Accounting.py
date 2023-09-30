import pandas as pd
import numpy as np
import os

path_Accounting = "D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Accounting"
os.chdir(path_Accounting)


tickers = pd.read_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Large Shareholders/Tickers.xlsx")
AvgLis = []

income_book_ratio = []
Intan_assets = []
book_asset = []
Debt_book_ratio = []
RD = []
Net_Income = []
Debt = []

for ticker in tickers['Ticker']:
    
    print("*************************************")
    print(ticker)
    
    #accounting = pd.read_excel('D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Accounting/{}.xlsx'.format(ticker))
    

    if(ticker == "کایگچ"):
        income_book_ratio.append(np.nan)
        Intan_assets.append(np.nan)
        book_asset.append(np.nan)
        Debt_book_ratio.append(np.nan)
        RD.append(np.nan) 
        Debt.append(np.nan)
        Net_Income.append(np.nan)
        continue
    
    accounting = pd.read_excel('D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Accounting/{}.xlsx'.format(ticker))
    
    if 'استهلاک' in list(accounting.columns):
        income_book_ratio.append(np.mean(accounting['درامد به ارزش دفتری'].dropna()))
        Intan_assets.append(np.mean(accounting['داریی های نامشهود '].dropna()))
        book_asset.append(np.mean(accounting['جمع داریی ها'].dropna()))
        Debt_book_ratio.append(np.mean(accounting['نسبت بدهی'].dropna()))
        RD.append(np.mean(accounting['استهلاک'].dropna()))
        Debt.append(np.mean(accounting['جمع بدهی ها '].dropna()))
        Net_Income.append(np.mean(accounting['درآمد خالص'].dropna()))
    else:

        income_book_ratio.append(np.mean(accounting['درامد به ارزش دفتری'].dropna()))
        Intan_assets.append(np.mean(accounting['داریی های نامشهود '].dropna()))
        book_asset.append(np.mean(accounting['جمع داریی ها'].dropna()))
        Debt_book_ratio.append(np.mean(accounting['نسبت بدهی'].dropna()))
        RD.append(np.mean(accounting['هزینه های عمومی و اداری '].dropna()))
        Debt.append(np.mean(accounting['جمع بدهی ها '].dropna()))
        Net_Income.append(np.mean(accounting['درآمد خالص'].dropna()))


dic = {'income_book_ratio': income_book_ratio, 'Intan_assets': Intan_assets , 'book_asset': book_asset , 'Debt_book_ratio': Debt_book_ratio , 'RD': RD , 'Debt': Debt , 'Net_Income': Net_Income}
    

Accounting = pd.DataFrame.from_dict(dic)
    
    

Accounting['RD'] =  -np.abs(Accounting['RD'])
Accounting.to_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Accounting.xlsx" , index = False)

path_TotalShare = "D:/Science/Economics/Term 6/Corporate Finance/Project/DataNew/Ownership"
path_price = "D:/Science/Economics/Term 6/Corporate Finance/Project/Price/Correct Name"

os.chdir(path_TotalShare)

Market_Cap = []

for idx in tickers['IDX']:
    
    print(idx)
    print("#####################################################")
    if(idx == 88):
        Market_Cap.append(np.nan) 
        continue
    where = []
    Total = pd.read_excel('D:/Science/Economics/Term 6/Corporate Finance/Project/DataNew/Ownership/{}.xlsx'.format(idx))
    price = pd.read_excel('D:/Science/Economics/Term 6/Corporate Finance/Project/Price/Correct Name/{}.xlsx'.format(idx))
    Total['total shares'].fillna(method='ffill', inplace=True)
    Total['total shares'].fillna(method='backfill', inplace=True)
    
    Dates = list(Total['Date'])
    
    for date in Dates:
        where.extend(price[price['<DTYYYYMMDD>'] == date].index.tolist())
    price = list(price['<CLOSE>'].iloc[where])
    total = list(Total['total shares'])[-len(price):]
    print(len(price))
    print(len(total))
    
    lis = []
    for i in range(len(total)):
        lis.append(total[i]*price[i])
    Market_Cap.append(np.mean(lis))
    
    
    
Accounting['Market_Cap'] = Market_Cap
Accounting['AQ'] = (Accounting['Market_Cap']/Accounting['book_asset'])/10**6


Accounting.to_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Accounting.xlsx" , index = False)
HHI = pd.read_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Industry_Value.xlsx")
Data = pd.read_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Data.xlsx")
Data = pd.merge(Data, HHI, how="left", on=['Industry Code'])

Data.to_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Data.xlsx" , index = False)