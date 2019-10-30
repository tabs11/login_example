import pandas as pd
import numpy as np
import glob
import datetime as dt
import re
def update_priority(path,company,prio_report):
    sheets=[]
    cis=pd.DataFrame()
    prio=pd.DataFrame()
    final=pd.DataFrame()
    path=path='/Users/numartin/Desktop/Priorities/'
    for j in range(len(glob.glob(path+'/*'))):
        for k in range(len(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names)):
            sheets.append(pd.read_excel(glob.glob(path+'*')[j],pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names[k]))     
    
    for j in range(len(sheets)):
        sheets[j].rename(columns=lambda x: re.sub('[^A-Za-z0-9]+', ' ', x), inplace=True)
        sheets[j].rename(columns=lambda x: x.strip(), inplace=True)
        sheets[j].dropna(axis=0,how='all',inplace=True)
    ###first overview prints####
        if (sheets[j].columns.str.contains('CI N',case=False).any()):
            cis=sheets[j][['CI Name','Site','Priority','Tier 1']][sheets[j]['Tier 1']=='RAN']
        if (sheets[j].columns.str.contains('VIP',case=False).any()):
            prio=sheets[j]
            prio['New Priority']='PRIORITY_1'
            prio.rename(columns={prio.columns[0]:'VIP Site'},inplace=True)
    if ((np.shape(cis)[0]>0) & (np.shape(prio)[0]>0)):
        cis['VIP Site']=cis['Site'].str.split("-", n = 1, expand = True)[1]
        cis['Priority']=cis['Priority'].str.upper()
    #    
    #    #list_prios=prio['New Priority'].unique()
        outer=cis.merge(prio,on='VIP Site',how='outer',indicator=True)
        inner=outer.copy()
        inner=inner[inner['_merge']=='both']
        inner['VIP Site']='Yes'
        left=outer.copy()
        left=left[left['_merge']=='left_only']
        left['VIP Site']='No'
    #    
        left['New Priority'] = np.where(left['Priority'] == 'PRIORITY_1', 'PRIORITY_5', left['New Priority'])
    #    #for i in range(len(list_prios)):
    #    #    left['New Priority'] = np.where(left['Priority'] == list_prios[i], 'PRIORITY_5', left['New Priority'])
        left['New Priority'].fillna(left['Priority'], inplace=True)
        final=pd.concat([inner,left],axis=0)
        final.rename(columns={'_merge':'match','Priority': 'Old Priority'},inplace=True)
        with pd.ExcelWriter(prio_report + company + '_priority_update_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
            final.to_excel(writer, 'priority_change',index=False)
    else:
        None
