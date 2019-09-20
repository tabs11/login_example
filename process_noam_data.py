import pandas as pd
import numpy as np
import os
import re
from openpyxl import load_workbook
import itertools
import datetime as dt
import glob
def noam_files(file_path,company,NOAM_report):

	noam_sites = load_workbook('CMDB_templates/Location_original.xlsm',read_only=False, keep_vba=True)
	sheets_sites = noam_sites.sheetnames
	w_sheet1_sites=noam_sites[sheets_sites[1]]
	w_sheet2_sites=noam_sites[sheets_sites[2]]
	w_sheet3_sites=noam_sites[sheets_sites[3]]
	w_sheet4_sites=noam_sites[sheets_sites[4]]
	w_sheet5_sites=noam_sites[sheets_sites[5]]
	##noam cis
	noam_cis = load_workbook('CMDB_templates/Transactional_CI_original.xlsm',read_only=False, keep_vba=True)
	sheets_cis = noam_cis.sheetnames
	w_sheet2_cis=noam_cis[sheets_cis[2]]
	##files to upload
	sheets=[]
	cis=[]
	sites=[]
	for j in range(len(glob.glob(file_path+'/*'))):
		for k in range(len(pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names)):
			sheets.append(pd.read_excel(glob.glob(file_path+'*')[j],pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names[k]))     
	
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
			w_sheet1_sites['A' +str(4+i)]=sites[0].filter(regex=re.compile('SITE N',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet4_sites['A' +str(4+i)]=sites[0].filter(regex=re.compile('SITE N',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet5_sites['A' +str(4+i)]=sites[0].filter(regex=re.compile('SITE N',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet5_sites['B' +str(4+i)]=sites[0].filter(regex=re.compile('ALIAS',re.IGNORECASE)).iloc[:,0].values[i]
			##street

			w_sheet1_sites['B' +str(4+i)]=sites[0].filter(regex=re.compile('STREET',re.IGNORECASE)).iloc[:,0].values[i]
			
			##country
			w_sheet1_sites['C' +str(4+i)]=sites[0].filter(regex=re.compile('COUNTRY',re.IGNORECASE)).iloc[:,0].values[i]
			##city
			w_sheet1_sites['E' +str(4+i)]=sites[0].filter(regex=re.compile('CITY',re.IGNORECASE)).iloc[:,0].values[i]
			##Location ID
			w_sheet1_sites['O' +str(4+i)]=sites[0].filter(regex=re.compile('Location',re.IGNORECASE)).iloc[:,0].values[i]
			##Description
			w_sheet1_sites['P' +str(4+i)]=sites[0].filter(regex=re.compile('Descr',re.IGNORECASE)).iloc[:,0].values[i]

			w_sheet3_sites['D' +str(4+i)]=sites[0].filter(regex=re.compile('Descr',re.IGNORECASE)).iloc[:,0].values[i]

			##Additional Sites Details
			w_sheet1_sites['Q' +str(4+i)]=sites[0].filter(regex=re.compile('Additional',re.IGNORECASE)).iloc[:,0].values[i]

			
			#status
			if (sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0]=='').any():
				w_sheet1_sites['R' +str(4+i)]='Enabled'
				w_sheet2_sites['C' +str(4+i)]='Enabled'
				
				w_sheet3_sites['E' +str(4+i)]='Enabled'
				w_sheet4_sites['E' +str(4+i)]='Enabled'
			else:
				w_sheet1_sites['R' +str(4+i)]=sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0].values[i]
				w_sheet2_sites['C' +str(4+i)]=sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0].values[i]

				w_sheet3_sites['E' +str(4+i)]=sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0].values[i]

				w_sheet4_sites['E' +str(4+i)]=sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0].values[i]

			
			#Latitude
			w_sheet1_sites['S' +str(4+i)]=sites[0].filter(regex=re.compile('LAT',re.IGNORECASE)).iloc[:,0].values[i]

			#LON
			w_sheet1_sites['T' +str(4+i)]=sites[0].filter(regex=re.compile('LON',re.IGNORECASE)).iloc[:,0].values[i]
			#Maintenance cirlce name
			w_sheet1_sites['Y' +str(4+i)]=sites[0].filter(regex=re.compile('CIRCLE',re.IGNORECASE)).iloc[:,0].values[i]
			#SITE TYPE
			w_sheet1_sites['Z' +str(4+i)]=sites[0].filter(regex=re.compile('TYPE',re.IGNORECASE)).iloc[:,0].values[i]
			##reg
			w_sheet2_sites['A' +str(4+i)]=sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet3_sites['C' +str(4+i)]=sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0].values[i]

			w_sheet4_sites['C' +str(4+i)]=sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0].values[i]
			#company
			w_sheet2_sites['B' +str(4+i)]=sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet3_sites['B' +str(4+i)]=sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet4_sites['B' +str(4+i)]=sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0].values[i]
			#Site Group
			w_sheet3_sites['A' +str(4+i)]=sites[0].filter(regex=re.compile('GROUP',re.IGNORECASE)).iloc[:,0].values[i]
			w_sheet4_sites['D' +str(4+i)]=sites[0].filter(regex=re.compile('GROUP',re.IGNORECASE)).iloc[:,0].values[i]
	
		noam_sites.save(filename = NOAM_report + company +'_sites_Noam.xlsm')
	else:
		None
	if len(cis)>0:
		cis=cis[0]
		cis.fillna('',inplace=True)
		for i in range(np.shape(cis)[0]):
			#static fields
			w_sheet2_cis['DJ' +str(4+i)]='migrator'
			w_sheet2_cis['DN' +str(4+i)]='Computer System'
			w_sheet2_cis['DT' +str(4+i)]='Computer System'
			w_sheet2_cis['DU' +str(4+i)]='BMC_COMPUTERSYSTEM'

			#CI type
			w_sheet2_cis['CL' +str(4+i)]='BMC_COMPUTERSYSTEM'
			##Product type
			w_sheet2_cis['CD' +str(4+i)]='Hardware'
			####Supported
			w_sheet2_cis['AC' +str(4+i)]='Yes'
			###CI Name
			w_sheet2_cis['D' +str(4+i)]=cis.filter(regex=re.compile('CI NAME',re.IGNORECASE)).iloc[:,0].values[i]

			##CI ID
			#w_sheet2_cis.write(4+i,11,cis.filter(regex=re.compile('CI ID',re.IGNORECASE)).iloc[:,0][i])
			###CI Description
			w_sheet2_cis['G' +str(4+i)]=cis.filter(regex=re.compile('CI Description',re.IGNORECASE)).iloc[:,0].values[i]

			####Status
			if (cis.filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0]=='').any():
				w_sheet2_cis['AP' +str(4+i)]='Deployed'
			else:
				w_sheet2_cis['AP' +str(4+i)]=cis.filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0].values[i]
			

			####Tier 1
			w_sheet2_cis['B' +str(4+i)]=cis['Tier 1'][i]

			####Tier 2
			w_sheet2_cis['H' +str(4+i)]=cis['Tier 2'][i]

			####Tier 3
			w_sheet2_cis['J' +str(4+i)]=cis['Tier 3'][i]

			###Product Name
			#w_sheet2_cis.write(4+i,13,cis['Product Name+'][i])
			w_sheet2_cis['N' +str(4+i)]=cis.filter(regex=re.compile('Product Name',re.IGNORECASE)).iloc[:,0].values[i]

			####Manufacturer
			w_sheet2_cis['V' +str(4+i)]=cis['Manufacturer'][i]

			####System Role
			w_sheet2_cis['AN' +str(4+i)]=cis['System Role'][i]

			###Priority
			if (cis['Priority']=='').any():
				w_sheet2_cis['BB' +str(4+i)]='PRIORITY_5'
			else:
				w_sheet2_cis['BB' +str(4+i)]=cis['Priority'][i]

			###Additional Information
			w_sheet2_cis['BE' +str(4+i)]=cis['Additional Information'][i]
			##
			###Region
			w_sheet2_cis['AL' +str(4+i)]=cis['Region'][i]
			##
			###Site Group
			w_sheet2_cis['AR' +str(4+i)]=cis['Site Group'][i]
			##
			###Site
			w_sheet2_cis['AT' +str(4+i)]=cis[cis.columns[~cis.columns.str.contains('Group',case=False)].tolist()].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0].values[i]
			##
			###Tag Number
			w_sheet2_cis['S' +str(4+i)]=cis['Tag Number'][i]

			##
			###Model Version
			###w_sheet2_cis.write(4+i,19,cis[0].filter(regex=re.compile('Version',re.IGNORECASE)).iloc[:,0][i])
			##
			###DNS
			w_sheet2_cis['AJ' +str(4+i)]=cis['DNS Host Name'][i]

			##
			###Domain
			w_sheet2_cis['AQ' +str(4+i)]=cis['Domain'][i]


			##
			###num cells
			###w_sheet2_cis.write(4+i,126,cis[0].filter(regex=re.compile('cells',re.IGNORECASE)).iloc[:,0][i])
		noam_cis.save(filename = NOAM_report + company +'_CIs_Noam.xlsm')
		#wb_cis.save(NOAM_report + company +'_cis_Noam.xls')
	else:
		None
#import xlrd
#from xlutils.copy import copy
#def noam_files(file_path,company,NOAM_report):
#	noam_sites = xlrd.open_workbook('CMDB_templates/Location_original.xls',formatting_info=True)
#	wb_sites=copy(noam_sites)
#	w_sheet1_sites=wb_sites.get_sheet(1)
#	w_sheet2_sites=wb_sites.get_sheet(2)
#	w_sheet3_sites=wb_sites.get_sheet(3)
#	w_sheet4_sites=wb_sites.get_sheet(4)
#	w_sheet5_sites=wb_sites.get_sheet(5)
#	##noam cis
#	noam_cis = xlrd.open_workbook('CMDB_templates/Transactional_CI_original.xls',formatting_info=True)
#	wb_cis=copy(noam_cis)
#	w_sheet2_cis=wb_cis.get_sheet(2)
#	##files to upload
#	sheets=[]
#	cis=[]
#	sites=[]
#	for j in range(len(glob.glob(file_path+'/*'))):
#		for k in range(len(pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names)):
#			sheets.append(pd.read_excel(glob.glob(file_path+'*')[j],pd.ExcelFile(glob.glob(file_path+'*')[j]).sheet_names[k]))     
#
#		#sheets[j].fillna('',inplace=True)
#	
#	for j in range(len(sheets)):
#		if (~sheets[j].columns.str.contains('CI N',case=False).any()) & (sheets[j].columns.str.contains('SITE N|SITE+|SITE*',case=False).any()):
#			sites.append(sheets[j])
#		else:
#			None
#		#CIS
#		if sheets[j].columns.str.contains('CI N',case=False).any():
#			cis.append(sheets[j])
#		else:
#			None
#
#
#	if len(sites)>0:
#		sites[0].fillna('',inplace=True)
#		for i in range(np.shape(sites[0])[0]):
#		#site
#			w_sheet1_sites.write(3+i,0,sites[0].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet4_sites.write(3+i,0,sites[0].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet5_sites.write(3+i,0,sites[0].filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet5_sites.write(3+i,1,sites[0].filter(regex=re.compile('ALIAS',re.IGNORECASE)).iloc[:,0][i])
#			##street
#			w_sheet1_sites.write(3+i,1,sites[0].filter(regex=re.compile('STREET',re.IGNORECASE)).iloc[:,0][i])
#			##country
#			w_sheet1_sites.write(3+i,2,sites[0].filter(regex=re.compile('COUNTRY',re.IGNORECASE)).iloc[:,0][i])
#			##city
#			w_sheet1_sites.write(3+i,4,sites[0].filter(regex=re.compile('CITY',re.IGNORECASE)).iloc[:,0][i])
#			##Location ID
#			w_sheet1_sites.write(3+i,14,sites[0].filter(regex=re.compile('Location',re.IGNORECASE)).iloc[:,0][i])
#			
#			##Description
#			w_sheet1_sites.write(3+i,15,sites[0].filter(regex=re.compile('Descr',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet3_sites.write(3+i,3,sites[0].filter(regex=re.compile('Descr',re.IGNORECASE)).iloc[:,0][i])
#			
#			##Additional Sites Details
#			w_sheet1_sites.write(3+i,16,sites[0].filter(regex=re.compile('Additional',re.IGNORECASE)).iloc[:,0][i])
#			
#			#status
#			if (sites[0]['Status*']=='').any():
#				w_sheet1_sites.write(3+i,17,'Enabled')
#				w_sheet2_sites.write(3+i,2,'Enabled')
#				w_sheet3_sites.write(3+i,4,'Enabled')
#				w_sheet4_sites.write(3+i,4,'Enabled')
#			else:
#				w_sheet1_sites.write(3+i,17,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
#				w_sheet2_sites.write(3+i,2,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
#				w_sheet3_sites.write(3+i,4,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
#				w_sheet4_sites.write(3+i,4,sites[0].filter(regex=re.compile('STATUS',re.IGNORECASE)).iloc[:,0][i])
#			
#			#Latitude
#			w_sheet1_sites.write(3+i,18,sites[0].filter(regex=re.compile('LAT',re.IGNORECASE)).iloc[:,0][i])
#			#Latitude
#			w_sheet1_sites.write(3+i,19,sites[0].filter(regex=re.compile('LON',re.IGNORECASE)).iloc[:,0][i])
#			
#			#Maintenance cirlce name
#			w_sheet1_sites.write(3+i,24,sites[0].filter(regex=re.compile('CIRCLE',re.IGNORECASE)).iloc[:,0][i])
#			
#			#SITE TYPE
#			w_sheet1_sites.write(3+i,25,sites[0].filter(regex=re.compile('TYPE',re.IGNORECASE)).iloc[:,0][i])
#			
#			##reg
#			w_sheet2_sites.write(3+i,0,sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet3_sites.write(3+i,2,sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet4_sites.write(3+i,2,sites[0].filter(regex=re.compile('REGION',re.IGNORECASE)).iloc[:,0][i])
#		
#			#company
#			w_sheet2_sites.write(3+i,1,sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet3_sites.write(3+i,1,sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet4_sites.write(3+i,1,sites[0].filter(regex=re.compile('COMPANY',re.IGNORECASE)).iloc[:,0][i])
#			
#			#Site Group
#			w_sheet3_sites.write(3+i,0,sites[0].filter(regex=re.compile('GROUP',re.IGNORECASE)).iloc[:,0][i])
#			w_sheet4_sites.write(3+i,3,sites[0].filter(regex=re.compile('GROUP',re.IGNORECASE)).iloc[:,0][i])
#			
#	
#		wb_sites.save(NOAM_report + company +'_sites_Noam.xls')
#	else:
#		None
#	if len(cis)>0:
#		cis=cis[0]
#		cis.fillna('',inplace=True)
#		for i in range(np.shape(cis)[0]):
#			#CI type
#			w_sheet2_cis.write(3+i,113,'migrator')
#			w_sheet2_cis.write(3+i,117,'Computer System')
#			w_sheet2_cis.write(3+i,123,'Computer System')
#			w_sheet2_cis.write(3+i,124,'BMC_COMPUTERSYSTEM')
#			#CI type
#			w_sheet2_cis.write(3+i,89,'BMC_COMPUTERSYSTEM')
#			##Product type
#			w_sheet2_cis.write(3+i,81,'Hardware')
#			###CI Name
#			w_sheet2_cis.write(3+i,3,cis['CI Name*'][i])
#			##CI ID
#			##w_sheet2_cis.write(3+i,11,cis[0].filter(regex=re.compile('CI ID',re.IGNORECASE)).iloc[:,0][i])
#			###CI Description
#			w_sheet2_cis.write(3+i,6,cis['CI Description'][i])
#			####Status
#			if (cis['Status*']=='').any():
#				w_sheet2_cis.write(3+i,41,'Deployed')
#			else:
#				w_sheet2_cis.write(3+i,41,cis['Status*'][i])
#			####Supported
#			w_sheet2_cis.write(3+i,28,'Yes')
#			####Tier 1
#			w_sheet2_cis.write(3+i,1,cis['Tier 1'][i])
#			####Tier 2
#			w_sheet2_cis.write(3+i,7,cis['Tier 2'][i])
#			####Tier 3
#			w_sheet2_cis.write(3+i,9,cis['Tier 3'][i])
#			###Product Name
#			w_sheet2_cis.write(3+i,13,cis['Product Name+'][i])
#			####Manufacturer
#			w_sheet2_cis.write(3+i,21,cis['Manufacturer'][i])
#			####System Role
#			w_sheet2_cis.write(3+i,39,cis['System Role'][i])
#			###Priority
#			if (cis['Priority']=='').any():
#				w_sheet2_cis.write(3+i,53,'PRIORITY_5')   
#			else:
#				w_sheet2_cis.write(3+i,53,cis['Priority'][i])
#			###Additional Information
#			###w_sheet2_cis.write(3+i,56,cis[0].filter(regex=re.compile('Additional',re.IGNORECASE)).iloc[:,0][i])
#			##
#			###Region
#			w_sheet2_cis.write(3+i,37,cis['Region'][i])
#			##
#			###Site Group
#			w_sheet2_cis.write(3+i,43,cis['Site Group'][i])
#			##
#			###Site
#			w_sheet2_cis.write(3+i,45,cis['Site+'][i])
#			##
#			###Tag Number
#			w_sheet2_cis.write(3+i,18,cis['Tag Number'][i])
#			##
#			###Model Version
#			###w_sheet2_cis.write(3+i,19,cis[0].filter(regex=re.compile('Version',re.IGNORECASE)).iloc[:,0][i])
#			##
#			###DNS
#			w_sheet2_cis.write(3+i,35,cis['DNS Host Name'][i])
#			##
#			###Domain
#			w_sheet2_cis.write(3+i,42,cis['Domain'][i])
#			##
#			###num cells
#			###w_sheet2_cis.write(3+i,126,cis[0].filter(regex=re.compile('cells',re.IGNORECASE)).iloc[:,0][i])
#
#		wb_cis.save(NOAM_report + company +'_cis_Noam.xls')
#	else:
#		None
#
