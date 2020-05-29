import pandas as pd
import numpy as np
import glob
def split_zte_file(path,report):
    df = pd.read_csv(glob.glob(path + '/*')[0],sep=';',low_memory=False)
    df.fillna('',inplace=True)
    #zte=zte[zte['vendorName']=='ZTE']
    zte=df[['UserLabel','NEType','NE','vendorName']]
    #zte=zte[~zte['vendorUnitFamilyType'].str.contains('BBU')]
    zte['UserLabel_old']=zte['UserLabel']
    #zte=zte[~zte['UserLabel'].str.contains("_LCK")]
    zte['NE'] = zte['NE'].str.split(',', 1).str[0]
    zte['NE'] = zte['NE'].str.split('=', 1).str[1]
    #zte=df[['UserLabel','NEType','Subnetwork']]
    zte.drop_duplicates(inplace=True)
    zte_BTS=zte[zte['NEType']=='SDR']
    ##2G
    zte_BSC=zte[zte['NEType']=='BSC']
    zte_BSC.rename(columns={'UserLabel':'BSC_name'},inplace=True)
    ####
    #zte_BCF=zte_BTS[(zte_BTS['NE'].astype(float)>1) & (zte_BTS['NE'].astype(float)<44)]
    zte_BCF=zte_BTS[zte_BTS['UserLabel'].astype(str).apply(lambda x:not(x.startswith('e',0) or x.startswith('u',0)))]
    zte_BCF.rename(columns={'UserLabel':'BCF_name'},inplace=True)
    zte_BCF['BCF_name']=zte_BCF['BCF_name'].str.split('_', 1).str[0]
    zte_BCF.drop_duplicates(inplace=True)
    zte_2G=zte_BSC.merge(zte_BCF,on='NE',how='outer',indicator=True)
    zte_2G.drop_duplicates(inplace=True)
    zte_2G.to_csv(report + '2G_ZTE.csv',index=False)
    #3G
    zte_RNC=zte[zte['NEType']=='RNC']
    zte_RNC.rename(columns={'UserLabel':'RNC_name'},inplace=True)
    ####
    #zte_WBTS=zte_BTS[(zte_BTS['NE'].astype(float)>100) & (zte_BTS['NE'].astype(float)<144)]
    zte_WBTS=zte_BTS[zte_BTS['UserLabel'].astype(str).apply(lambda x:x.startswith('u',0))]
    zte_WBTS.rename(columns={'UserLabel':'WBTS_name'},inplace=True)
    zte_WBTS['WBTS_name']=zte_WBTS['WBTS_name'].str.split('_', 1).str[0]
    zte_WBTS.drop_duplicates(inplace=True)
    zte_3G=zte_WBTS.merge(zte_RNC,on='NE',how='outer',indicator=True)
    zte_3G.drop_duplicates(inplace=True)
    zte_3G.to_csv(report + '3G_ZTE.csv',index=False)

    #4G
    #zte_4G=zte_BTS[(zte_BTS['NE'].astype(float)>43) & (zte_BTS['NE'].astype(float)<101) | (zte_BTS['NE'].astype(float)>143)]
    zte_4G=zte_BTS[zte_BTS['UserLabel'].astype(str).apply(lambda x:x.startswith('e',0))]
    zte_4G.rename(columns={'UserLabel':'LNBTS_name'},inplace=True)
    zte_4G['LNBTS_name']=zte_4G['LNBTS_name'].str.split('_', 1).str[0]
    zte_4G.drop_duplicates(inplace=True)
    zte_4G.to_csv(report + '4G_ZTE.csv',index=False)