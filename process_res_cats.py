# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import re
import glob

# =============================================================================
# Create DataFrame
# =============================================================================
def rescats_files(file_path,company,res_cats_report):
    res_cats=load_workbook('CMDB_templates/Generic_Catalog.xlsm',read_only=False, keep_vba=True)
    sheets_sites = res_cats.sheetnames
    w_sheet1_res_cats=res_cats[sheets_sites[1]]
    w_sheet2_res_cats=res_cats[sheets_sites[2]]
    w_sheet3_res_cats=res_cats[sheets_sites[3]]
    df2=pd.read_excel(glob.glob(file_path+'*')[0])
    df2['Tier3'][df2['Tier3'].str.contains('None')]='- None -'
    for k in range(len(df2.columns)):
        df2.iloc[:,k]=df2.iloc[:,k].apply(lambda x: x.strip() if type(x)==str else x)
    for i in range(df2.shape[0]):      
        w_sheet1_res_cats['A' +str(4+i)]='Resolution Category'
        w_sheet1_res_cats['B' +str(4+i)]=df2.filter(regex=re.compile('Tier1',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet1_res_cats['C' +str(4+i)]=df2.filter(regex=re.compile('Tier2',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet1_res_cats['D' +str(4+i)]=df2.filter(regex=re.compile('Tier3',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet1_res_cats['O' +str(4+i)]='Enabled'
        w_sheet2_res_cats['A' +str(4+i)]=df2.filter(regex=re.compile('Company',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet2_res_cats['B' +str(4+i)]='Enabled'
        w_sheet2_res_cats['C' +str(4+i)]='Resolution Category'
        w_sheet2_res_cats['D' +str(4+i)]=df2.filter(regex=re.compile('Tier1',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet2_res_cats['E' +str(4+i)]=df2.filter(regex=re.compile('Tier2',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet2_res_cats['F' +str(4+i)]=df2.filter(regex=re.compile('Tier3',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet2_res_cats['L' +str(4+i)]=0
        w_sheet2_res_cats['M' +str(4+i)]=0
        w_sheet2_res_cats['N' +str(4+i)]=0
        w_sheet2_res_cats['O' +str(4+i)]=0
        w_sheet2_res_cats['P' +str(4+i)]=0
        w_sheet3_res_cats['A' +str(4+i)]='Resolution Category'
        w_sheet3_res_cats['B' +str(4+i)]=df2.filter(regex=re.compile('Tier1',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet3_res_cats['C' +str(4+i)]=df2.filter(regex=re.compile('Tier2',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet3_res_cats['D' +str(4+i)]=df2.filter(regex=re.compile('Tier3',re.IGNORECASE)).iloc[:,0].values[i]
        w_sheet3_res_cats['F' +str(4+i)]='Enabled'
    res_cats.save(filename = res_cats_report + company + '_res_cats_example.xlsm') 

		
