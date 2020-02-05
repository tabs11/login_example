import win32com.client as com
import re
import pandas as pd
import datetime as dt
import csv
import numpy as np

import pythoncom,threading, time


def sites_cis_report(company,site_report):
    pythoncom.CoInitialize()
    # Get instance
    outlook = com.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Create id
    #xl_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, outlook)
    
    folder = outlook.Folders[0]
    #inbox = folder.Folders[1]
    inbox = outlook.Folders("nuno_david.martins.ext@nokia.com").Folders[1]
    print(inbox)
    subject='Andreas Schilling team - ITSM CMDB PROD system - Export (Status: deployed) for customer: ' + company
    a=[]
    for message in inbox.Items:
        if subject in message.Subject and message.UnRead == True:#and 'chilling' in str(message.Sender)
            a.append(message.Body)
    if len(a)==0:
        print('No emails related')
    else:
        #print('emails')
        s=pd.Series(a[0].split('\n'))[1:]
        report=s.str.split(';',expand=True)
        report=pd.DataFrame(report.rename(columns=report.iloc[0]))
        report.drop(columns='\r',inplace=True)
        index_site=report.index[report['Company'] == 'Site data:\r'].tolist()[0]
        sites_history=report[index_site:]
        sites_history.columns=sites_history.iloc[0,:]
        sites_history=sites_history[sites_history['Site Name'].notnull()][1:][sites_history.columns.dropna()]
        sites_history.columns=[s.replace('\r', '') for s in sites_history.columns.tolist()]
        #sites_history=sites_history[sites_history['PrimAlias']=='0']
        #sites_history.drop('PrimAlias',axis=1,inplace=True)
        #sites_history['PrimAlias']=pd.Series(np.where(sites_history['PrimAlias']=='0','Yes','No'))
        sites_history.rename(columns={'Maint Circle Name':'Maintenance Circle Name'},inplace=True)
        sites_history['Additional Site Details']=''
        sites_history['Status']='Enabled'
        sites_history.rename(columns=lambda x: re.sub('[^A-Za-z0-9]+', ' ', x), inplace=True)
        sites_history.rename(columns=lambda x: x.strip(), inplace=True)
        sites_history.to_csv(site_report + company + '_Sites_report.csv',sep=';',mode='w',index=False)
        index_cis=report[:index_site]
        cis_history=index_cis[index_cis[' CI_Name'].notnull()][1:][index_cis.columns.dropna()]
        cis_history.drop(['TagNumber','NbrCells'],axis=1,inplace=True)
        cis_history.rename(columns={
            'ProdCat2':'Tag Number','ProdCat3':'NbrCells',
            'Manufacturer_Name':'Manufacturer',
            'Prodcat1':'Tier 1', 
            'Prodcat2': 'Tier 2',
            'Prodcat3': 'Tier 3',
            'SiteGroup':'Site Group',\
            'SystemRole':'System Role',
            cis_history.columns[11]:'Additional Information'},inplace=True)
        cis_history['Status']='Deployed'
        cis_history['CI ID']=''
        #cis_history['CI type']=''
        #cis_history['Product type']=''
        #cis_history['Supported']=''
        cis_history['Model Version']='' 
        #cis_history['Additional Information']=''
        cis_history.rename(columns=lambda x: re.sub('[^A-Za-z0-9]+', ' ', x), inplace=True)
        cis_history.rename(columns=lambda x: x.strip(), inplace=True)
        cis_history.to_csv(site_report + company + '_CIs_report.csv',sep=';',mode='w',index=False)

    #cis_history=pd.concat([pd.Series([1,2,3]).rename('A'),pd.Series([1,2,3]).rename('B')],axis=1)
    #sites_history=pd.concat([pd.Series([4,5,6]).rename('A'),pd.Series([4,5,6]).rename('B')],axis=1)
        #with pd.ExcelWriter(site_report + company +'_sites_cis_report.xlsx',engine='xlsxwriter') as writer:
        #    sites_history.to_excel(writer,'sites',index=False)
        #    cis_history.to_excel(writer,'cis',index=False)
        #    writer.save()