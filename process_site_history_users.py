import numpy as np
import pandas as pd
import datetime as dt
from pandas import DataFrame
import sqlite3
from tabulate import tabulate


def sites_cis_report(user,company,site_report):
    #cmdb_owners=pd.read_excel('CMDB_templates/cmdb Owners_full_list.xlsx')
    #if ((cmdb_owners['Login ID']==user) & (cmdb_owners['Company']==company)).any() | (user in ['numartin','paulof','mccavitt','paagrawa'] or company=='Dummy Company'):

    conn = sqlite3.connect('CMDB_inventory/CMDB_data_PROD.db')  # You can create a new database by changing the name within the quotes
    c = conn.cursor() # The database will be saved in the location where your 'py' file is saved
    quer_sites=r"""SELECT DISTINCT * FROM SITES WHERE Company='"""+company+"""'"""
    quer_cis=r"""SELECT DISTINCT * FROM CIS WHERE Company='"""+company+"""'"""
    ###get sites
    c.execute(quer_sites)
    sites_itsm = DataFrame(c.fetchall(), columns=['Company','Site Name','Site Alias','Description','Region','Site Group','Street','Country','City','Latitude','Longitude','Maintenance Circle Name','Site Type','Location ID','PrimAlias','Additional Site Details','Status','Date'])
    sites_itsm.to_csv(site_report + company + '_Sites_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
    ###get cis
    c.execute(quer_cis)
    cis_itsm = DataFrame(c.fetchall(), columns=['Company','CI Name','Site','Region','Site Group','CI Description','DNS Host Name','System Role','Product Name','Tier 1','Tier 2','Tier 3','Manufacturer','Model Version','Additional Information','Tag Number','CI ID','NbrCells','Domain','Status','Reconciliation Identity','Priority','Date'])
    cis_itsm.to_csv(site_report + company + '_CIs_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
    if len(sites_itsm)>0:
        sites_size=len(sites_itsm['Site Name'].drop_duplicates())
    else:
        sites_size='No Sites found in inventory'
            
    if len(cis_itsm)>0:
        cis_size=len(cis_itsm)
    else:
        cis_size='No CIs found in inventory'
    #####
    ###count all sites per company
    #c.execute('''SELECT Company,COUNT(Company),COUNT(DISTINCT([Site Name]))
    #FROM SITES
    #GROUP BY Company''')
    #all_sites=pd.DataFrame(c.fetchall())
    ####count all cis per company
    #c.execute('''SELECT Company, COUNT(Company)
    #FROM CIS
    #GROUP BY Company''')
    #all_cis=pd.concat([pd.DataFrame(c.fetchall()),pd.DataFrame([pd.Series(['T-Mobile US','records not found'])])],axis=0)
    conn.commit()
    conn.close()
    counts_cmb=pd.concat([pd.Series(['Sites','CIs']).rename('LEVEL'),pd.Series([sites_size,cis_size]).rename('COUNT'),pd.Series([np.unique(sites_itsm['Date'])[0],np.unique(cis_itsm['Date'])[0]]).rename('Last Date Report')],axis=1)
    print('',tabulate(counts_cmb,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(site_report +'SQLDB_CMDB.txt','a',encoding='utf8'))
    
    ####all cmdb
    #all_cmdb=all_sites.merge(all_cis,on=0,how='inner').rename(columns={0:'Company','1_x':'Total Sites (including Alias)',2:'Total Sites','1_y':'CIS Total'})
    #print('Count of all CMDB',tabulate(all_cmdb,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(site_report +'SQLDB_CMDB.txt','a',encoding='utf8'))


    #else:
    #    print('','Not authorized user for this CMDB','',sep='\n',file=open(site_report +'SQLDB_CMDB.txt','a',encoding='utf8'))

    
