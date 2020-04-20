import pandas as pd
import itertools
import os
import glob
import re
import numpy as np
from tabulate import tabulate
import datetime as dt
import sqlite3
from pandas import DataFrame
import process_cmdb_inventory
######process file
def process_rap_file(path,company,report):
    global cis_itsm
    cis_itsm=pd.DataFrame()
   ####read_file
    column_names=['RA Bulk Upload Template']
    rap=pd.read_excel(glob.glob(path+'*')[0],list(set(pd.ExcelFile(glob.glob(path+'*')[0]).sheet_names).intersection(column_names))[0],header=1)
    rap.rename(columns=lambda x: x.strip(), inplace=True)
    rap.dropna(axis=0,how='all',inplace=True)
    rap=rap[rap['Description']!='INSERT LINES ABOVE THIS ROW']
    rap.rename(columns={rap.filter(regex=re.compile('CI',re.IGNORECASE)).columns[0]:'CI Name'},inplace=True)
    ##check blanks
    blank_cases=[]

    for i in range(len(rap.columns)):
        blank_find=rap.iloc[:,i][rap.iloc[:,i].astype(str).apply(lambda x: x[0].isspace() or x[len(x)-1].isspace())]
        blank_cases.append(pd.concat([pd.Series(blank_find.name).rename('FIELD'),pd.Series(len(blank_find)).rename('COUNT')],axis=1))
        rap.iloc[:,i]=rap.iloc[:,i].apply(lambda x: x.strip() if type(x)==str else x)
    blanks=pd.concat(blank_cases)
    blanks=blanks[blanks['COUNT']>0]
    if len(blanks)>0:
        print('','Fields with blanks spaces: (Blanks Auto Removed) '.upper(),'-'*len('Fields with blanks spaces: (Blanks Auto Removed) '),tabulate(blanks,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
    else:
        None
    #check null values
    null_columns=rap[rap.columns[rap.isnull().any()]] 
    field_type=[['Yes']*12,['No']*3]
    filed_type=pd.concat([pd.Series(rap.columns.tolist()).rename('FIELD'),pd.Series(list(itertools.chain(*field_type))).rename('Mandatory')],axis=1)
    if np.shape(null_columns.isnull().sum())[0]>0:
        null_fields=pd.DataFrame(null_columns.isnull().sum())
        null_fields.reset_index(level=0, inplace=True)
        null_fields.rename(columns={'index':'FIELD',0:'COUNT'},inplace=True)
        null_fields=null_fields.merge(filed_type,on='FIELD',how='inner')
        print('','Fields with Null Values: '.upper(),'-'*len('Fields with Null Values:'),tabulate(null_fields,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
    
    else:
        None
    categories=[
        ########fix frequencyNumber####
        [0,1,2,3],
        ['Days','Weeks','Months','Years'],
        ['1 Day','2 Days','5 Days','10 Days'],
        ['1 Day','2 Days','3 Days','1 Week','2 Weeks','1 Month'],
        ['On Site FSO','Remote TSO'],
        ['No Impact','Service Impact']
    ]
    categoric_fields=rap[['Frequency','Recommended Frequency','RA Trigger','Window','Intervention Type','Service Impact']]
    issues=[]
    for i in range(len(categories)):
        if len(list(set(categoric_fields.iloc[:,i]) - set(categories[i])))>0:
            categoric_fields.iloc[:,i][categoric_fields.iloc[:,i]=='']="NULL"
            issues.append(categoric_fields.iloc[:,i][categoric_fields.iloc[:,i].isin(list(set(categoric_fields.iloc[:,i]) - set(categories[i])))])
            print('','WRONG ' +issues[i].name+' values ('+ str(categories[i])+')','-'*len('WRONG ' +issues[i].name+' values ('+ str(categories[i])+')'),tabulate(pd.DataFrame(issues[i]),headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
        else:    
            None

    ###remove duplicated lines
    rap.drop_duplicates(inplace=True)

    if rap['CI Name'].isnull().all():
        print('NO CIs to check')
    else:
        duplicate_cis=[]
        count_issues=[]
        ##get cmdb CIs
        process_cmdb_inventory.call_cmdb_inventory(company,report)
        ###correlation between rap cis and cmdb cis
        cis=rap[['CI Name']].merge(cis_itsm,on='CI Name',how='outer',indicator=True)
        ###cis not foundin CMDB
        cis_not_existing=pd.DataFrame(cis[cis['_merge']=='left_only'].dropna()['CI Name'])
        cis_not_existing_count=pd.Series(len(cis_not_existing))
        #print(pd.concat([pd.Series('CI Name not existing in CMDB:'.upper()),cis_not_existing_count],axis=1))
        count_issues.append(pd.concat([pd.Series('CI Name not existing in CMDB:'.upper()),cis_not_existing_count],axis=1))
        
        ####duplicated CIs
        filtered_cis=rap.filter(regex=re.compile('CI N',re.IGNORECASE))
        dup_cis=pd.DataFrame(rap[filtered_cis.duplicated(keep=False)].drop_duplicates()).sort_values(['CI Name'])
        duplicate_cis.append(dup_cis.groupby(filtered_cis.columns.values[0]).size().reset_index(name='counts'))
        if len(dup_cis)>0:
            dup_cis_count=pd.Series(len(duplicate_cis[0]))
            count_issues.append(pd.concat([pd.Series('Duplicated CI Names (excluding duplicate rows):'.upper()),dup_cis_count],axis=1))
            #print('','Duplicated CI Names (excluding duplicate rows): '.upper(),'-'*len('Duplicated CI Names (excluding duplicate rows): '.upper()),tabulate(dup_cis_count,headers=['COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding=
            #print('','Duplicated CI Names (excluding duplicate rows): '.upper(),'-'*len('Duplicated CI Names (excluding duplicate rows):'),tabulate(dup_sites_cis.set_index(dup_sites_cis.columns[0]),headers=['COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
        else:
            None
        #print('','CI Name not existing in CMDB:','-'*len('CI Name not existing in CMDB:'),tabulate(cis,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
        if len(count_issues)>0:             
            print('','ISSUES IN CIS FIELD:','-'*len('ISSUES IN CIS FIELD:'),tabulate(pd.concat(count_issues),headers=['ISSUE','COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
        else:
            None
#####save file        
    with pd.ExcelWriter(report + company + '_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
        rap.to_excel(writer, 'RA Bulk Upload Template',index=False)
        if np.shape(cis_not_existing)[0]>0:
            cis_not_existing.to_excel(writer, 'Non existing CI',index=False)
        else:
            None
        if np.shape(dup_cis)[0]>0:
            dup_cis.to_excel(writer, 'Duplicate CIs',index=False)
        else:
            None
        writer.save()#




#def process_file(path,company,report):
#    ####read_file
#    sheets=[]
#    column_names=['RA Bulk Upload Template']
#    sheets.append(pd.read_excel(glob.glob(path+'*')[0],list(set(pd.ExcelFile(glob.glob(path+'*')[0]).sheet_names).intersection(column_names))[0],header=1))
#    rap=sheets[0]
#    rap.rename(columns=lambda x: x.strip(), inplace=True)
#    rap.dropna(axis=0,how='all',inplace=True)
#    rap=rap[rap['Description']!='INSERT LINES ABOVE THIS ROW']
#    ##check blanks
#    blank_cases=[]
#
#    for i in range(len(rap.columns)):
#        blank_find=rap.iloc[:,i][rap.iloc[:,i].astype(str).apply(lambda x: x[0].isspace() or x[len(x)-1].isspace())]
#        blank_cases.append(pd.concat([pd.Series(blank_find.name).rename('FIELD'),pd.Series(len(blank_find)).rename('COUNT')],axis=1))
#        rap.iloc[:,i]=rap.iloc[:,i].apply(lambda x: x.strip() if type(x)==str else x)
#    blanks=pd.concat(blank_cases)
#    blanks=blanks[blanks['COUNT']>0]
#    print(blanks)
#    if len(blanks)>0:
#    	print('','Fields with blanks spaces: (Blanks Auto Removed) '.upper(),'-'*len('Fields with blanks spaces: (Blanks Auto Removed) '),tabulate(blanks,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
#    else:
#    	None
#    #check null values
#    null_columns=rap[rap.columns[rap.isnull().any()]] 
#    field_type=[['Yes']*12,['No']*3]
#    filed_type=pd.concat([pd.Series(rap.columns.tolist()).rename('FIELD'),pd.Series(list(itertools.chain(*field_type))).rename('Mandatory')],axis=1)
#    if np.shape(null_columns.isnull().sum())[0]>0:
#        null_fields=pd.DataFrame(null_columns.isnull().sum())
#        null_fields.reset_index(level=0, inplace=True)
#        null_fields.rename(columns={'index':'FIELD',0:'COUNT'},inplace=True)
#        null_fields=null_fields.merge(filed_type,on='FIELD',how='inner')
#        print('','Fields with Null Values: '.upper(),'-'*len('Fields with Null Values:'),tabulate(null_fields,headers="keys",tablefmt="fancy_grid",showindex=False),'',)
#    
#    else:
#        None
#    categories=[
#        [0,1,2,3],
#        ['Days','Weeks','Months','Years'],
#        ['1 Day','2 Days','5 Days','10 Days'],
#        ['1 Day','2 Days','3 Days','1 Week','2 Weeks','1 Month'],
#        ['On Site FSO','Remote TSO'],
#        ['No Impact','Service Impact']
#    ]
#    categoric_fields=rap[['Frequency','Recommended Frequency','RA Trigger','Window','Intervention Type','Service Impact']]
#    issues=[]
#    for i in range(len(categories)):
#        if len(list(set(categoric_fields.iloc[:,i]) - set(categories[i])))>0:
#            categoric_fields.iloc[:,i][categoric_fields.iloc[:,i]=='']="NULL"
#            issues.append(categoric_fields.iloc[:,i][categoric_fields.iloc[:,i].isin(list(set(categoric_fields.iloc[:,i]) - set(categories[i])))])
#            print('','WRONG ' +issues[i].name+' values ('+ str(categories[i])+')','-'*len('WRONG ' +issues[i].name+' values ('+ str(categories[i])+')'),tabulate(pd.DataFrame(issues[i]),headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
#        else:    
#            None
#
#    #####save file        
#    with pd.ExcelWriter(report + company + '_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
#        rap.to_excel(writer, 'RA Bulk Upload Template',index=False)
#        writer.save()#
#
#