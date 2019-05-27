import pandas as pd
import numpy as np
import re
import os
from difflib import get_close_matches
from openpyxl import load_workbook
import itertools
import datetime as dt
import glob
import xlwt
import xlrd
from xlutils.copy import copy

def noam_files(file_path,company,NOAM_report):
	###noamsites
	noam_sites = xlrd.open_workbook('CMDB_templates/Location_updated.xls',formatting_info=True)
	wb_sites=copy(noam_sites)
	w_sheet1_sites=wb_sites.get_sheet(1)
	w_sheet2_sites=wb_sites.get_sheet(2)
	w_sheet3_sites=wb_sites.get_sheet(3)
	w_sheet4_sites=wb_sites.get_sheet(4)
	w_sheet5_sites=wb_sites.get_sheet(5)
	##noam cis
	noam_cis = xlrd.open_workbook('CMDB_templates/Transactional_CI_updated.xls',formatting_info=True)
	wb_cis=copy(noam_cis)
	w_sheet2_cis=wb_cis.get_sheet(2)
	##files to upload
	sheets=[]
	cis=[]
	sites=[]
	for j in range(len(glob.glob(file_path+'/*'))):
		for k in range(len(pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names)):
			sheets.append(pd.read_excel(glob.glob(file_path+'*')[j],pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names[k]))     

		#sheets[j].fillna('',inplace=True)
	
	for j in range(len(sheets)):
		if (~sheets[j].columns.str.contains('CI N',case=False).any()) & (sheets[j].columns.str.contains('SITE N|SITE+|SITE*',case=False).any()):
			sites.append(sheets[j])
		else:
			None
		#CIS
		if sheets[j].columns.str.contains('CI N',case=False).any():
			cis.append(sheets[j])
		else:
			None


	if len(sites)>0:
		sites[0].fillna('',inplace=True)
		for i in range(np.shape(sites[0])[0]):
		#site
			w_sheet1_sites.write(3+i,0,sites[0].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0][i])
			w_sheet4_sites.write(3+i,0,sites[0].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0][i])
			w_sheet5_sites.write(3+i,0,sites[0].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0][i])
			w_sheet5_sites.write(3+i,1,sites[0].filter(regex=re.compile('ALIAS',re.IGNORECASE)).iloc[:,0][i])
			##street
			w_sheet1_sites.write(3+i,1,sites[0].filter(regex=re.compile('STREET',re.IGNORECASE)).iloc[:,0][i])
			##country
			w_sheet1_sites.write(3+i,2,sites[0].filter(regex=re.compile('COUNTRY',re.IGNORECASE)).iloc[:,0][i])
			##city
			w_sheet1_sites.write(3+i,4,sites[0].filter(regex=re.compile('CITY',re.IGNORECASE)).iloc[:,0][i])
			##Location ID
			w_sheet1_sites.write(3+i,14,sites[0].filter(regex=re.compile('Location',re.IGNORECASE)).iloc[:,0][i])
			
			##Description
			w_sheet1_sites.write(3+i,15,sites[0].filter(regex=re.compile('Descr',re.IGNORECASE)).iloc[:,0][i])
			w_sheet3_sites.write(3+i,3,sites[0].filter(regex=re.compile('Descr',re.IGNORECASE)).iloc[:,0][i])
			
			##Additional Sites Details
			w_sheet1_sites.write(3+i,16,sites[0].filter(regex=re.compile('Additional',re.IGNORECASE)).iloc[:,0][i])
			
			#status
			w_sheet1_sites.write(3+i,17,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
			w_sheet2_sites.write(3+i,2,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
			w_sheet3_sites.write(3+i,4,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
			w_sheet4_sites.write(3+i,4,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
			
			#Latitude
			w_sheet1_sites.write(3+i,18,sites[0].filter(regex=re.compile('LAT',re.IGNORECASE)).iloc[:,0][i])
			#Latitude
			w_sheet1_sites.write(3+i,19,sites[0].filter(regex=re.compile('LON',re.IGNORECASE)).iloc[:,0][i])
			
			#Maintenance cirlce name
			w_sheet1_sites.write(3+i,24,sites[0].filter(regex=re.compile('CIRCLE',re.IGNORECASE)).iloc[:,0][i])
			
			#SITE TYPE
			w_sheet1_sites.write(3+i,25,sites[0].filter(regex=re.compile('TYPE',re.IGNORECASE)).iloc[:,0][i])
			
			##reg
			w_sheet2_sites.write(3+i,0,sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0][i])
			w_sheet3_sites.write(3+i,2,sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0][i])
			w_sheet4_sites.write(3+i,2,sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0][i])
		
			#company
			w_sheet2_sites.write(3+i,1,sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0][i])
			w_sheet3_sites.write(3+i,1,sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0][i])
			w_sheet4_sites.write(3+i,1,sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0][i])
			
			#Site Group
			w_sheet3_sites.write(3+i,0,sites[0].filter(regex=re.compile('GROUP',re.IGNORECASE)).iloc[:,0][i])
			w_sheet4_sites.write(3+i,3,sites[0].filter(regex=re.compile('GROUP',re.IGNORECASE)).iloc[:,0][i])
	
		wb_sites.save(NOAM_report + company +'_sites_Noam.xls')
	else:
		None
	if len(cis)>0:
		cis[0].fillna('',inplace=True)
		for i in range(np.shape(cis[0])[0]):
			#CI type
			w_sheet2_cis.write(3+i,113,'migrator')
			w_sheet2_cis.write(3+i,117,'Computer System')
			w_sheet2_cis.write(3+i,123,'Computer System')
			w_sheet2_cis.write(3+i,124,'BMC_COMPUTERSYSTEM')

			cis[0]['CI type']='BMC_COMPUTERSYSTEM'
			w_sheet2_cis.write(3+i,89,cis[0].filter(regex=re.compile('CI TYPE',re.IGNORECASE)).iloc[:,0][i])
			#Product type
			cis[0]['Product type']='Hardware'
			w_sheet2_cis.write(3+i,81,cis[0].filter(regex=re.compile('Product Type',re.IGNORECASE)).iloc[:,0][i])
			##CI Name
			w_sheet2_cis.write(3+i,3,cis[0].filter(regex=re.compile('CI NAME',re.IGNORECASE)).iloc[:,0][i])
			#CI ID
			#w_sheet2_cis.write(3+i,11,cis[0].filter(regex=re.compile('CI ID',re.IGNORECASE)).iloc[:,0][i])
			##CI Description
			w_sheet2_cis.write(3+i,6,cis[0].filter(regex=re.compile('CI DESCR',re.IGNORECASE)).iloc[:,0][i])
			##Status
			if (cis[0]['Status*']=='').any():
				cis[0]['Status*']='Deployed'
			else:
				None
			w_sheet2_cis.write(3+i,41,cis[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
			##Supported
			if (cis[0]['Supported']=='').any():
				cis[0]['Supported']='Yes'    
			else:
				None
			w_sheet2_cis.write(3+i,28,cis[0].filter(regex=re.compile('SUPPORTED',re.IGNORECASE)).iloc[:,0][i])
			##Tier 1
			w_sheet2_cis.write(3+i,1,cis[0].filter(regex=re.compile('Tier 1',re.IGNORECASE)).iloc[:,0][i])
			##Tier 2
			w_sheet2_cis.write(3+i,7,cis[0].filter(regex=re.compile('Tier 2',re.IGNORECASE)).iloc[:,0][i])
			##Tier 3
			w_sheet2_cis.write(3+i,9,cis[0].filter(regex=re.compile('Tier 3',re.IGNORECASE)).iloc[:,0][i])
			##Product Name
			w_sheet2_cis.write(3+i,13,cis[0].filter(regex=re.compile('Product N',re.IGNORECASE)).iloc[:,0][i])
			##Manufacturer
			w_sheet2_cis.write(3+i,21,cis[0].filter(regex=re.compile('Manuf',re.IGNORECASE)).iloc[:,0][i])
			##System Role
			w_sheet2_cis.write(3+i,39,cis[0].filter(regex=re.compile('Role',re.IGNORECASE)).iloc[:,0][i])
			#Priority
			if (cis[0]['Priority']=='').any():
				cis[0]['Priority']='PRIORITY_5'    
			else:
				None
			w_sheet2_cis.write(3+i,53,cis[0].filter(regex=re.compile('Priority',re.IGNORECASE)).iloc[:,0][i])
			#Additional Information
			#w_sheet2_cis.write(3+i,56,cis[0].filter(regex=re.compile('Additional',re.IGNORECASE)).iloc[:,0][i])
			
			#Region
			w_sheet2_cis.write(3+i,37,cis[0].filter(regex=re.compile('Region',re.IGNORECASE)).iloc[:,0][i])
			
			#Site Group
			w_sheet2_cis.write(3+i,43,cis[0].filter(regex=re.compile('Group',re.IGNORECASE)).iloc[:,0][i])
			
			#Site
			w_sheet2_cis.write(3+i,45,cis[0].filter(regex=re.compile('Site',re.IGNORECASE)).iloc[:,0][i])
			
			#Tag Number
			w_sheet2_cis.write(3+i,18,cis[0].filter(regex=re.compile('TAG Number',re.IGNORECASE)).iloc[:,0][i])
			
			#Model Version
			#w_sheet2_cis.write(3+i,19,cis[0].filter(regex=re.compile('Version',re.IGNORECASE)).iloc[:,0][i])
			
			#DNS
			w_sheet2_cis.write(3+i,35,cis[0].filter(regex=re.compile('DNS',re.IGNORECASE)).iloc[:,0][i])
			
			#Domain
			w_sheet2_cis.write(3+i,42,cis[0].filter(regex=re.compile('Domain',re.IGNORECASE)).iloc[:,0][i])
			
			#num cells
			#w_sheet2_cis.write(3+i,126,cis[0].filter(regex=re.compile('cells',re.IGNORECASE)).iloc[:,0][i])
		
		wb_cis.save(NOAM_report + company +'_cis_Noam.xls')
	else:
		None