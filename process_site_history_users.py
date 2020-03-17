import pandas as pd
import datetime as dt
from pandas import DataFrame
import sqlite3


def sites_cis_report(company,report):
    conn = sqlite3.connect('/Users/numartin/Desktop/DVT_PROPOSAL/CMDB_inventory/CMDB_data.db')  # You can create a new database by changing the name within the quotes
    c = conn.cursor() # The database will be saved in the location where your 'py' file is saved
    quer_sites=r"""SELECT DISTINCT * FROM SITES WHERE Company='"""+company+"""'"""
    quer_cis=r"""SELECT DISTINCT * FROM CIS WHERE Company='"""+company+"""'"""
    ###get sites
    c.execute(quer_sites)
    sites_itsm = DataFrame(c.fetchall(), columns=['Company','Site Name','Region','Site Group'])
    sites_itsm.to_csv(report + company + '_Sites_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
    ###get cis
    c.execute(quer_cis)
    cis_itsm = DataFrame(c.fetchall(), columns=['Company','CI Name'])
    cis_itsm.to_csv(report + company + '_CIs_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
    conn.commit()
    conn.close()


