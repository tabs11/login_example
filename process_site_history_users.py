import pandas as pd
import datetime as dt
from pandas import DataFrame
import sqlite3
from tabulate import tabulate


def sites_cis_report(company,site_report):
    conn = sqlite3.connect('CMDB_inventory/CMDB_data.db')  # You can create a new database by changing the name within the quotes
    c = conn.cursor() # The database will be saved in the location where your 'py' file is saved
    quer_sites=r"""SELECT DISTINCT * FROM SITES WHERE Company='"""+company+"""'"""
    quer_cis=r"""SELECT DISTINCT * FROM CIS WHERE Company='"""+company+"""'"""
    ###get sites
    c.execute(quer_sites)
    sites_itsm = DataFrame(c.fetchall(), columns=['Company','Site Name','Region','Site Group'])
    sites_itsm.to_csv(site_report + company + '_Sites_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
    ###get cis
    c.execute(quer_cis)
    cis_itsm = DataFrame(c.fetchall(), columns=['Company','CI Name'])
    cis_itsm.to_csv(site_report + company + '_CIs_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
    if len(sites_itsm)>0:
        sites_size=len(sites_itsm)
    else:
        sites_size='No Sites found in inventory'
            
    if len(cis_itsm)>0:
        cis_size=len(cis_itsm)
    else:
        cis_size='No CIs found in inventory'
    
    counts_cmb=pd.concat([pd.Series(['Sites','CIs']).rename('LEVEL'),pd.Series([sites_size,cis_size]).rename('COUNT')],axis=1)
    print('',tabulate(counts_cmb,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(site_report +'SQLDB_CMDB.txt','a',encoding='utf8'))
    conn.commit()
    conn.close()