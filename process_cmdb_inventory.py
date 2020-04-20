from lib import *
import process_rap
#import process_common_val


#def read_file(path):
#    column_names=['sites','cis','Site Data Template','CI Data Template Comp Syst']
#    count_issues=[]
#    for j in range(len(glob.glob(path+'/*'))):
#        if glob.glob(path+'/*')[j].endswith(('.xls','.xlsx')):
#            for k in range(len(list(set(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names).intersection(column_names)))):
#                process_common_val.sheets.append(pd.read_excel(glob.glob(path+'*')[j],list(set(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names).intersection(column_names))[k]))
#            #files=pd.read_excel(glob.glob(path+'*')[j],sheet_name=None)
#            #for frame in files.keys():
#            #   process_data.sheets.append(files[frame])
#        elif glob.glob(path+'/*')[j].endswith('.csv'):
#            process_common_val.sheets.append(pd.read_csv(glob.glob(path+'*')[j],sep=";",encoding='ISO-8859-1'))
#        else:
#            None


def call_cmdb_inventory(company,report):
    ####cmdb inventory

    conn = sqlite3.connect('CMDB_inventory/CMDB_data.db')  # You can create a new database by changing the name within the quotes
    c = conn.cursor() # The database will be saved in the location where your 'py' file is saved
    quer_sites=r"""SELECT DISTINCT * FROM SITES WHERE Company='"""+company+"""'"""
    quer_cis=r"""SELECT DISTINCT * FROM CIS WHERE Company='"""+company+"""'"""
    ###get sites
    c.execute(quer_sites)
    process_rap.sites_itsm = DataFrame(c.fetchall(), columns=['Company','Site Name','Site Alias','Description','Region','Site Group','Street','Country','City','Latitude','Longitude','Maintenance Circle Name','Site Type','Location ID','PrimAlias','Additional Site Details','Status','Date'])
    process_rap.sites_itsm.drop(['Company'],axis=1,inplace=True)
    ###get cis
    c.execute(quer_cis)
    process_rap.cis_itsm = DataFrame(c.fetchall(), columns=['Company','CI Name','Site','Region','Site Group','CI Description','DNS Host Name','System Role','Product Name','Tier 1','Tier 2','Tier 3','Manufacturer','Model Version','Additional Information','Tag Number','CI ID','NbrCells','Domain','Status','Reconciliation Identity','Priority','Date'])
    process_rap.cis_itsm.drop(['Company'],axis=1,inplace=True)
    conn.close()
    