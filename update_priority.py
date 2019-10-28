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
    for j in range(len(glob.glob(path+'/*'))):
        for k in range(len(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names)):
            sheets.append(pd.read_excel(glob.glob(path+'*')[j],pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names[k]))     

    for j in range(len(sheets)):
        sheets[j].rename(columns=lambda x: re.sub('[^A-Za-z0-9]+', ' ', x), inplace=True)
        sheets[j].rename(columns=lambda x: x.strip(), inplace=True)
        sheets[j].dropna(axis=0,how='all',inplace=True)
    ###first overview prints####
        if (sheets[j].columns.str.contains('CI N',case=False).any()):
            cis=sheets[j][['CI Name','Site','Priority']]
        if (sheets[j].columns.str.contains('VIP',case=False).any()):
            prio=sheets[j]
    if ((np.shape(cis)[0]>0) & (np.shape(prio)[0]>0)):
        cis['VIP locations']=cis['Site'].str.split("-", n = 1, expand = True)[1]
        cis['Priority']=cis['Priority'].str.upper()
        
        list_prios=prio['New Priority'].unique()
        outer=cis.merge(prio,on='VIP locations',how='outer',indicator=True)
        inner=outer[outer['_merge']=='both']
        inner.drop(['VIP locations'],axis=1, inplace=True)
        left=outer[outer['_merge']=='left_only']
        left.drop(['VIP locations'],axis=1, inplace=True)
        for i in range(len(list_prios)):
            left['New Priority'] = np.where(left['Priority'] == list_prios[i], 'PRIORITY_5', left['New Priority'])
        left['New Priority'].fillna(left['Priority'], inplace=True)
        final=pd.concat([inner,left],axis=0)
        final.rename(columns={'_merge':'match','Priority': 'Old Priority'},inplace=True)
        print('Success')
        with pd.ExcelWriter(prio_report + company + '_priority_update_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
            final.to_excel(writer, 'priority_change',index=False)
    else:
        None

