import glob
import re
import pandas as pd
import datetime as dt
import csv
import numpy as np

import sqlite3
from tabulate import tabulate

def cmdb_update(path,company):
    sheets=[]
    for j in range(len(glob.glob(path+'/*'))):
        if glob.glob(path+'/*')[j].endswith('.csv'):
            sheets.append(pd.read_csv(glob.glob(path+'*')[j],sep=";",encoding='ISO-8859-1'))
        else:
            None  
    conn = sqlite3.connect('CMDB_inventory/CMDB_data.db')
    c = conn.cursor()
    conn.commit()
    quer_sites=r"""DELETE FROM SITES WHERE Company='"""+company+"""'"""
    quer_cis=r"""DELETE FROM CIS WHERE Company='"""+company+"""'"""
    c.execute(quer_cis)
    c.execute(quer_sites)
    for j in range(len(sheets)):
        if (~sheets[j].columns.str.contains('CI N',case=False).any()):
            sheets[j].to_sql('SITES',conn,if_exists='append', index = False)
        elif (sheets[j].columns.str.contains('CI N',case=False).any()):
            sheets[j].to_sql('CIS',conn,if_exists='append', index = False) 
        else:
            None
    
    conn.commit()#
    conn.close()
    #print('',tabulate(company + 'CMDB SUCCESSFULLY UPDATED','',sep='\n',file=open(site_report +'SQLDB_CMDB.txt','a',encoding='utf8'))

