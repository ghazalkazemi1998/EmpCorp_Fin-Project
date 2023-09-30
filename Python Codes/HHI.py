import numpy as np
import itertools
import pandas as pd
from scipy.sparse import csr_matrix , csc_matrix
from random import sample
from scipy.sparse.linalg import inv
import statsmodels.api as sm
from scipy.sparse import identity
from queue import Queue
import pandas as pd
from scipy.interpolate import interp1d
import matplotlib.pyplot as plt
from scipy import optimize 
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.datasets import load_iris
from sklearn import metrics
import seaborn as sn
from google.colab import drive
drive.mount('/content/drive')



Industry = pd.read_excel('drive/My Drive/ByIndustry.xlsx')

Industry_Value = Industry.sort_values(['ارزش'], ascending=False).groupby('صنعت').sum()
Industry_Value['index'] = [i for i in range(len(Industry_Value))]
Industry_Value = Industry_Value.reset_index().set_index('index')
Industry_Value['HHI'] = np.nan

Total_Industry = pd.merge(Industry,Industry_Value,on='صنعت',how='left')
Total_Industry['MarketShare'] = Total_Industry['ارزش_x'] / Total_Industry['ارزش_y']
Total_Industry['MSSqured'] = (Total_Industry['MarketShare']*100)**2

for industry in Industry_Value['صنعت']:
  Industry = Total_Industry[Total_Industry['صنعت'] == industry]
  if len(Industry) > 3:
    Industry_Value.loc[Industry_Value['صنعت'] == industry] = np.sum(Total_Industry['MSSqured'].iloc[0:3])
  else:
    Industry_Value.loc[Industry_Value['صنعت'] == industry] = np.sum(Total_Industry['MSSqured'].iloc[0:3])
    

