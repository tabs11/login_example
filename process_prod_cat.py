import glob
import numpy as np
import pandas as pd
import datetime as dt
from pandas import DataFrame
from tabulate import tabulate
import re
def prod_update(path):
    existing_prod_cat=pd.read_csv('Prod_Cats_V2/oneitsm_ProdCats_new_version.csv',sep=";")
    if glob.glob(path+'*')[0].endswith(('.xls','.xlsx')):
        new_prod_cat=pd.read_excel(glob.glob(path+'*')[0])
    elif glob.glob(TEMP_FOLDER+'*')[0].endswith('.csv'):
        new_prod_cat=pd.csv(glob.glob(path+'*')[0],sep=';')
    
    new_prod_cat.rename(columns=lambda x: re.sub('[^A-Za-z0-9]+', ' ', x), inplace=True)
    new_prod_cat.rename(columns=lambda x: x.strip(), inplace=True)
    if (list(set(new_prod_cat.columns.tolist()))==list(set(existing_prod_cat.columns.tolist()))):
        existing_prod_cat=pd.concat([existing_prod_cat,new_prod_cat],axis=0).drop_duplicates()
        existing_prod_cat.to_csv('Prod_Cats_V2/oneitsm_ProdCats_new_version.csv',sep=';',mode='w',index=False)
    
    else:
        print('Check the column names:',existing_prod_cat.columns.tolist())

    