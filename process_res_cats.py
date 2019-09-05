# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import re
import glob

def op_res_cats_files(file_path,company,op_res_cats_report):

    ##Operational Category
    op_cats=load_workbook('CMDB_templates/OperationalCatalog.xlsm',read_only=False, keep_vba=True)
    sheets_ops = op_cats.sheetnames
    w_sheet_op_cats=op_cats[sheets_ops[1]]
    w_sheet1_op_cats=op_cats[sheets_ops[2]]
    #Resolution Category
    res_cats=load_workbook('CMDB_templates/Generic_Catalog.xlsm',read_only=False, keep_vba=True)
    sheets_res = res_cats.sheetnames
    w_sheet1_res_cats=res_cats[sheets_res[1]]
    w_sheet2_res_cats=res_cats[sheets_res[2]]
    w_sheet3_res_cats=res_cats[sheets_res[3]]
    sheets=[]
    res=pd.DataFrame()
    ops=pd.DataFrame()
    for j in range(len(glob.glob(file_path+'/*'))):
        for k in range(len(pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names)):
            sheets.append(pd.read_excel(glob.glob(file_path+'*')[j],pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names[k])) 
    
    for j in range(len(sheets)):
        #res
        if (sheets[j].columns.str.contains('ResCat',case=False).any()):
            res=sheets[j]
        else:
            None
        #ops
        if sheets[j].columns.str.contains('OpCat',case=False).any():
            ops=sheets[j]
        else:
            None
    
    if len(res)>0:
        res.fillna('',inplace=True)
        res['ResCat3'][res['ResCat3'].str.contains('None')]='- None -'
        for k in range(len(res.columns)):
            res.iloc[:,k]=res.iloc[:,k].apply(lambda x: x.strip() if type(x)==str else x)
        for i in range(res.shape[0]):      
            w_sheet1_res_cats['A' +str(4+i)]='Resolution Category'
            w_sheet1_res_cats['B' +str(4+i)]=res.filter(regex=re.compile('ResCat1',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet1_res_cats['C' +str(4+i)]=res.filter(regex=re.compile('ResCat2',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet1_res_cats['D' +str(4+i)]=res.filter(regex=re.compile('ResCat3',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet1_res_cats['O' +str(4+i)]='Enabled'
            w_sheet2_res_cats['A' +str(4+i)]=res.filter(regex=re.compile('Company',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet2_res_cats['B' +str(4+i)]='Enabled'
            w_sheet2_res_cats['C' +str(4+i)]='Resolution Category'
            w_sheet2_res_cats['D' +str(4+i)]=res.filter(regex=re.compile('ResCat1',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet2_res_cats['E' +str(4+i)]=res.filter(regex=re.compile('ResCat2',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet2_res_cats['F' +str(4+i)]=res.filter(regex=re.compile('ResCat3',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet2_res_cats['L' +str(4+i)]='Yes'
            w_sheet2_res_cats['M' +str(4+i)]='Yes'
            w_sheet2_res_cats['N' +str(4+i)]='Yes'
            w_sheet2_res_cats['O' +str(4+i)]='Yes'
            w_sheet2_res_cats['P' +str(4+i)]='Yes'
            w_sheet3_res_cats['A' +str(4+i)]='Resolution Category'
            w_sheet3_res_cats['B' +str(4+i)]=res.filter(regex=re.compile('ResCat1',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet3_res_cats['C' +str(4+i)]=res.filter(regex=re.compile('ResCat2',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet3_res_cats['D' +str(4+i)]=res.filter(regex=re.compile('ResCat3',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet3_res_cats['F' +str(4+i)]='Enabled'
        res_cats.save(filename = op_res_cats_report + company + '_res_cats_example.xlsm')
    if len(ops)>0:
        ops.fillna('',inplace=True)
        ops['inc_values'] = np.where(ops['Module'] == "Incident Management", 'Yes', None)
        ops['prb_values'] = np.where(ops['Module'] == "Problem Management", 'Yes', None)
        ops['chg_values'] = np.where(ops['Module'] == "Change Management", 'Yes', None)
        ops['conf_values'] = np.where(ops['Module'] == "Configuration/Asset Management", 'Yes', None)
        ops['rel_values'] = np.where(ops['Module'] == "Release Management", 'Yes', None)
    
        ops['OpCat3'][ops['OpCat3'].str.contains('None')]='- None -'
        for k in range(len(ops.columns)):
            ops.iloc[:,k]=ops.iloc[:,k].apply(lambda x: x.strip() if type(x)==str else x)
        for i in range(ops.shape[0]): 
            w_sheet_op_cats['A' +str(4+i)]=ops.filter(regex=re.compile('Company',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet_op_cats['B' +str(4+i)]=ops.filter(regex=re.compile('OpCat1',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet_op_cats['C' +str(4+i)]=ops.filter(regex=re.compile('OpCat2',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet_op_cats['D' +str(4+i)]=ops.filter(regex=re.compile('OpCat3',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet_op_cats['Q' +str(4+i)]='Enabled'
            w_sheet1_op_cats['A' +str(4+i)]=ops.filter(regex=re.compile('OpCat1',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet1_op_cats['B' +str(4+i)]=ops.filter(regex=re.compile('OpCat2',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet1_op_cats['C' +str(4+i)]=ops.filter(regex=re.compile('OpCat3',re.IGNORECASE)).iloc[:,0].values[i]
            w_sheet1_op_cats['F' +str(4+i)]='Enabled'
            w_sheet_op_cats['F' +str(4+i)]=ops['inc_values'].values[i]
            w_sheet_op_cats['G' +str(4+i)]=ops['inc_values'].values[i]
            w_sheet_op_cats['H' +str(4+i)]=ops['inc_values'].values[i]
            w_sheet_op_cats['I' +str(4+i)]=ops['inc_values'].values[i]
            w_sheet_op_cats['J' +str(4+i)]=ops['inc_values'].values[i]
            w_sheet_op_cats['K' +str(4+i)]=ops['prb_values'].values[i]
            w_sheet_op_cats['N' +str(4+i)]=ops['chg_values'].values[i]
            w_sheet_op_cats['L' +str(4+i)]=ops['conf_values'].values[i]
            w_sheet_op_cats['R' +str(4+i)]=ops['rel_values'].values[i]
        op_cats.save(filename = op_res_cats_report + company + '_op_cats_example.xlsm')
# =============================================================================
# Create DataFrame
# =============================================================================
#def rescats_files(file_path,company,res_cats_report):
#    res_cats=load_workbook('CMDB_templates/Generic_Catalog.xlsm',read_only=False, keep_vba=True)
#    sheets_sites = res_cats.sheetnames
#    w_sheet1_res_cats=res_cats[sheets_sites[1]]
#    w_sheet2_res_cats=res_cats[sheets_sites[2]]
#    w_sheet3_res_cats=res_cats[sheets_sites[3]]
#    df2=pd.read_excel(glob.glob(file_path+'*')[0])
#    df2['Tier3'][df2['Tier3'].str.contains('None')]='- None -'
#    for k in range(len(df2.columns)):
#        df2.iloc[:,k]=df2.iloc[:,k].apply(lambda x: x.strip() if type(x)==str else x)
#    for i in range(df2.shape[0]):      
#        w_sheet1_res_cats['A' +str(4+i)]='Resolution Category'
#        w_sheet1_res_cats['B' +str(4+i)]=df2.filter(regex=re.compile('Tier1',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet1_res_cats['C' +str(4+i)]=df2.filter(regex=re.compile('Tier2',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet1_res_cats['D' +str(4+i)]=df2.filter(regex=re.compile('Tier3',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet1_res_cats['O' +str(4+i)]='Enabled'
#        w_sheet2_res_cats['A' +str(4+i)]=df2.filter(regex=re.compile('Company',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet2_res_cats['B' +str(4+i)]='Enabled'
#        w_sheet2_res_cats['C' +str(4+i)]='Resolution Category'
#        w_sheet2_res_cats['D' +str(4+i)]=df2.filter(regex=re.compile('Tier1',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet2_res_cats['E' +str(4+i)]=df2.filter(regex=re.compile('Tier2',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet2_res_cats['F' +str(4+i)]=df2.filter(regex=re.compile('Tier3',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet2_res_cats['L' +str(4+i)]=0
#        w_sheet2_res_cats['M' +str(4+i)]=0
#        w_sheet2_res_cats['N' +str(4+i)]=0
#        w_sheet2_res_cats['O' +str(4+i)]=0
#        w_sheet2_res_cats['P' +str(4+i)]=0
#        w_sheet3_res_cats['A' +str(4+i)]='Resolution Category'
#        w_sheet3_res_cats['B' +str(4+i)]=df2.filter(regex=re.compile('Tier1',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet3_res_cats['C' +str(4+i)]=df2.filter(regex=re.compile('Tier2',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet3_res_cats['D' +str(4+i)]=df2.filter(regex=re.compile('Tier3',re.IGNORECASE)).iloc[:,0].values[i]
#        w_sheet3_res_cats['F' +str(4+i)]='Enabled'
#    res_cats.save(filename = res_cats_report + company + '_res_cats_example.xlsm') 
#
#		
#
