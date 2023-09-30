import pandas as pd
import numpy as np
import os

path_TotalShare = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/Total Share"
path_Ownership = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/Ownership"
path_Haghighi = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/Haghighi"
path_HaghighiPrime = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/HaghighiPrime"
path_OwnershipPrime = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/OwnershipPrime"
path_HaghighiPercent = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/HaghighiPercent"
path_OwnershipPercent = "D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/OwnershipPercent"
os.chdir(path_TotalShare)

for file in os.listdir():
    Total = pd.read_excel(file)
    
    Total['total shares'].fillna(method='ffill', inplace=True)
    Total['total shares'].fillna(method='backfill', inplace=True)
    file_path = f"{path_Ownership}/{file}"
    Ownership = pd.read_excel(file_path)
    file_path = f"{path_Haghighi}/{file}"
    Haghighi = pd.read_excel(file_path)
    Ownership['total share'] = Total['total shares']
    Haghighi['total share'] = Total['total shares']
    print(file)
    
    Haghighi.to_excel(f"{path_HaghighiPrime}/{file}", index = False)
    Ownership.to_excel(f"{path_OwnershipPrime}/{file}", index = False)
    

os.chdir(path_HaghighiPrime)

for file in os.listdir():
    
    haghighi = pd.read_excel(file)
    print(haghighi['total share'].isnull().sum())
    
#     haghighi['total share'].fillna(method='ffill', inplace=True)
#     print(haghighi['total share'].isnull().sum())
#     #Haghighi.fillna(0 , inplace = True)
#     Haghighi.to_excel(f"{path_HaghighiPrime}/{file}", index = False)
    
#     file_path = f"{path_OwnershipPrime}/{file}"
#     Ownership = pd.read_excel(file_path)
    
#     Ownership['total share'].fillna(method='ffill', inplace=True)
#     #Ownership.fillna(0 , inplace = True)
#     Ownership.to_excel(f"{path_OwnershipPrime}/{file}", index = False)
#     print(file)
    
    

for file in os.listdir():
    haghighi = pd.read_excel(file)
    haghighi = haghighi.iloc[:,1:]
    
    print(file)
    print("####################################################")
    columns = list(haghighi.columns)[1:-1]
    
    print(haghighi.columns)
    for col in columns:
        haghighi[col] = haghighi[col]/haghighi['total share']
    
    print(haghighi.columns)
    
#     file_path = f"{path_OwnershipPrime}/{file}"
#     Ownership = pd.read_excel(file_path)
    
#     columns = list(Ownership.columns)[1:-1]
    
#     for col in columns:
#         Ownership[col] = Ownership[col]/Ownership['total share']

        
    haghighi.to_excel(f"{path_HaghighiPercent}/{file}" , index = False)
#    Ownership.to_excel(f"{path_OwnershipPercent}/{file}" , index = False)


for file in os.listdir():
    if file.endswith(".xlsx"):
        a = pd.read_excel(file)
        indices = np.where(a['name'] == 'شخص حقيقي')[0].tolist()
        ID_Haghighi = []
        for index in indices:
            ID_Haghighi.append(a.loc[index,'id'])
        file_path = f"{path_Ownership}/{file}"
        b = pd.read_excel(file_path)
        
        NewDataFrame = pd.DataFrame()
        NewDataFrame['Date'] = b['Date']
        
        for ID in ID_Haghighi:
            
            NewDataFrame[str(ID)] = b[str(ID)]
            
        List_Haghighi = NewDataFrame.columns.values.tolist()[1:]    
        
        NewDataFrame['Total'] = NewDataFrame[List_Haghighi].sum(axis = 1)
        NewDataFrame.to_excel(f"{path_Haghighi}/{file}")
        
        

os.chdir(path_OwnershipPercent)

tickers = pd.read_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Large Shareholders/Tickers.xlsx")
AvgLis = []

for idx in tickers['IDX']:
    
    print("*************************************")
    print(idx)
    
    ownership = pd.read_excel('D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/OwnershipPercent/{}.xlsx'.format(idx))
    
    lis = []
    for i in range(len(ownership)):
        b = list(ownership.iloc[i,1:-1].dropna())
        b.sort(reverse = True)
    
        if len(b) >= 5:
            lis.append(np.sum(b[:5]))
        elif len(b) > 0:
            lis.append(np.sum(b)) 
    
    AvgLis.append(np.mean(lis))
    
    print(np.mean(lis))
    
ticker = list(tickers['Ticker'])

dic = {'Ticker' : ticker , 'A5': AvgLis}

A5 = pd.DataFrame.from_dict(dic)

A5.to_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/A5.xlsx" , index = False)




os.chdir(path_HaghighiPercent)

tickers = pd.read_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Large Shareholders/Tickers.xlsx")
AvgLis = []

for idx in tickers['IDX']:
    
    print("*************************************")
    print(idx)
    
    Haghighi = pd.read_excel('D:/Science/Economics/Term 6/Corporate Finance/Project/Haghighi Data/HaghighiPercent/{}.xlsx'.format(idx)) 
    
    AvgLis.append(np.mean(Haghighi['Total']))
    
    print(np.mean(lis))
    
ticker = list(tickers['Ticker'])

dic = {'Ticker' : ticker , 'Mngt': AvgLis}

Mngt = pd.DataFrame.from_dict(dic)

Mngt.to_excel("D:/Science/Economics/Term 6/Corporate Finance/Project/Data/Mngt.xlsx" , index = False)





