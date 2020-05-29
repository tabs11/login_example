import xlsxwriter
import numpy as np
import pandas as pd
import re
import os
import glob
from difflib import get_close_matches
import itertools
import datetime as dt
from tabulate import tabulate
import sqlite3
from pandas import DataFrame
#import process_cmdb_inventory


def process_file(path,company,report,instance):
	sites_itsm=pd.DataFrame(columns=['Site Name','Region','Site Group'])
	cis_itsm=pd.DataFrame(columns=['CI Name'])
	all_sites=pd.DataFrame(columns=['Site','Region','Site Group'])

	######read files
	sheets=[]
	cis=pd.DataFrame()
	sites=pd.DataFrame()
	files=pd.DataFrame()
	column_names=['sites','cis','Site Data Template','CI Data Template Comp Syst']
	count_issues=[]
	for j in range(len(glob.glob(path+'/*'))):
		if glob.glob(path+'/*')[j].endswith(('.xls','.xlsx')):
			for k in range(len(list(set(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names).intersection(column_names)))):
				sheets.append(pd.read_excel(glob.glob(path+'*')[j],list(set(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names).intersection(column_names))[k]))
				#print(max(j,k))
			#files=pd.read_excel(glob.glob(path+'*')[j],sheet_name=None)
			#for frame in files.keys():
			#	sheets.append(files[frame])
		elif glob.glob(path+'/*')[j].endswith('.csv'):
			sheets.append(pd.read_csv(glob.glob(path+'*')[j],sep=";",encoding='ISO-8859-1'))
		else:
			None
	if len(sheets)==0:
		print('',sep='\n',file=open(report +'Mismatched_fields.txt','a',encoding='utf8'))
	else:
		##check correct fields
		
		sites_fields=['Company','Site Name','Site Alias','Description','Region','Site Group','Street','Country','City','Latitude','Longitude','Location ID','Additional Site Details','Maintenance Circle Name','Site Type','Status']
		cis_fields=['Company','CI Name','CI Description','Tag Number','System Role','Status','Priority','Additional Information','Tier 1','Tier 2','Tier 3','Product Name','Model Version','Manufacturer','Region','Site Group','Site','DNS Host Name','Domain','CI ID']
		fields=sites_fields+cis_fields
		char_num=[254,60,60,255,60,60,90,60,60,12,12,30,10,70,22,8,254,254,254,64,30,16,10,254,60,60,60,254,254,254,60,60,60,254,254,64]
		field_type=[['No'],['Yes']*2,['No'],['Yes']*2,['No'],['Yes']*2,['No']*6,['Yes'],['No'],['Yes'],['No']*3,['Yes']*2,['No'],['Yes']*4,['No'],['Yes']*4,['No']*3]
		field_type=list(itertools.chain(*field_type))     
		unmatched_fields=[]
		names=[]
		for j in range(len(sheets)):
			sheets[j].rename(columns=lambda x: re.sub('[^A-Za-z0-9]+', ' ', x), inplace=True)
			sheets[j].rename(columns=lambda x: x.strip(), inplace=True)
			sheets[j].dropna(axis=0,how='all',inplace=True)
		###first overview prints####
			if (~sheets[j].columns.str.contains('CI N',case=False).any()):
				names.append('Sites')
				unmatched_fields.append(list(set(sites_fields) - set(sheets[j])))
			if (sheets[j].columns.str.contains('CI N',case=False).any()):
				names.append('CIs')
				unmatched_fields.append(list(set(cis_fields) - set(sheets[j])))   
			else:
				None
		if len(list(itertools.chain(*unmatched_fields)))>0:
			unmatched=pd.DataFrame(pd.Series(list(itertools.chain(*unmatched_fields))).rename('FIELD'))
			#print('',unmatched.set_index('FIELD'),'',sep='\n',file=open(report +'Mismatched_fields.txt','a',encoding='utf8'))
			print('',tabulate(unmatched,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'Mismatched_fields.txt','a',encoding='utf8')) 
			#print('','No Warnings to display'.upper(),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
		else:
			####cmdb inventory
			conn = sqlite3.connect('CMDB_inventory/CMDB_data_'+instance+'.db')  # You can create a new database by changing the name within the quotes
			c = conn.cursor() # The database will be saved in the location where your 'py' file is saved
			quer_sites=r"""SELECT DISTINCT * FROM SITES WHERE Company='"""+company+"""'"""
			quer_cis=r"""SELECT DISTINCT * FROM CIS WHERE Company='"""+company+"""'"""
			###get sites
			c.execute(quer_sites)
			sites_itsm = DataFrame(c.fetchall(), columns=['Company','Site Name','Site Alias','Description','Region','Site Group','Street','Country','City','Latitude','Longitude','Maintenance Circle Name','Site Type','Location ID','PrimAlias','Additional Site Details','Status','Date'])
			
			###get cis
			c.execute(quer_cis)
			cis_itsm = DataFrame(c.fetchall(), columns=['Company','CI Name','Site','Region','Site Group','CI Description','DNS Host Name','System Role','Product Name','Tier 1','Tier 2','Tier 3','Manufacturer','Model Version','Additional Information','Tag Number','CI ID','NbrCells','Domain','Status','Reconciliation Identity','Priority','Date'])
			
			conn.close()
			###change in prod
			if (sites_itsm['Company']=='fDummy Company').all():
				sites_itsm=pd.DataFrame(columns=['Site Name','Region','Site Group'])
				cis_itsm=pd.DataFrame(columns=['CI Name'])
				print('','NO CMDB inventory choosen.','The validation was done without checking the CMDB.','',sep='\n',file=open(report +'summary_CMDB.txt','a',encoding='utf8'))      

			else:
				if (len(sites_itsm)>0):
				    sites_size=len(sites_itsm['Site Name'].drop_duplicates())
				else:
				    sites_size='No Sites found in inventory'
				
				if len(cis_itsm)>0:
				    cis_size=len(cis_itsm)
				else:
				    cis_size='No CIs found in inventory'
				
				#counts_cmb=pd.concat([pd.Series(['Sites','CIs']).rename('LEVEL'),pd.Series([sites_size,cis_size]).rename('COUNT')],axis=1)
				counts_cmb=pd.concat([pd.Series(['Sites','CIs']).rename('LEVEL'),pd.Series([sites_size,cis_size]).rename('COUNT'),pd.Series([np.unique(sites_itsm['Date'])[0],np.unique(cis_itsm['Date'])[0]]).rename('Last Date Report')],axis=1)
				sites_itsm=sites_itsm[['Site Name','Region','Site Group']]
				cis_itsm=cis_itsm[['CI Name']]
	
				print('',tabulate(counts_cmb,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'summary_CMDB.txt','a',encoding='utf8'))      


			#############
			all_fields=pd.concat([pd.Series(fields).rename('FIELD'),pd.Series(field_type).rename('MANDATORY'),pd.Series(char_num).rename('ALLOWED')],axis=1)
			common_fields=[]
			for j in range(len(sheets)):
				###first overview prints####
				print(sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
				print(sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				print(sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
				summary_list=[]
				if len(os.listdir(path))==1:
					summary_list.append(pd.concat([pd.Series('INPUT FILE NAME '),pd.Series(os.listdir(path)[0])],axis=1))

					#print('','INPUT FILE NAME: '+,'-'*len('INPUT FILE NAME:'),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
				else:
					summary_list.append(pd.concat([pd.Series('INPUT FILE NAME '),pd.Series(os.listdir(path)[j])],axis=1))
					#print('','INPUT FILE NAME: '+os.listdir(path)[j],'-'*len('INPUT FILE NAME:'),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
	
				summary_list.append(pd.concat([pd.Series('Number of records'.upper()),pd.Series(len(sheets[j]))],axis=1))

				#print('Number of records: '.upper()+str(np.shape(sheets[j])[0]),'-'*len('Number of records:'),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))					
				common_fields.append(all_fields.merge(pd.DataFrame(pd.Series(sheets[j].columns).rename('FIELD')),on='FIELD',how='inner').drop_duplicates())
				blank_cases=[]
				count_chars=[]
				#check null values
				
				null_columns=sheets[j][sheets[j].columns[sheets[j].isnull().any()]]  		
				###blanks and char num
				for i in range(len(sheets[j].columns)):
					blank_find=sheets[j].iloc[:,i][sheets[j].iloc[:,i].astype(str).apply(lambda x: x[0].isspace() or x[len(x)-1].isspace())]
					blank_cases.append(pd.concat([pd.Series(blank_find.name).rename('FIELD'),pd.Series(len(blank_find)).rename('COUNT')],axis=1))
					#if len(blank_find)>0:
					#blank_cases.append(blank_find.iloc[0])
	
					#else:
					#	blank_cases.append('None')
					####count_chars
					count_chars.append(sheets[j].iloc[:,i].apply(lambda x: x if pd.isnull(x) else (len(str(round(x,5))) if type(x)==float else len(str(x)))).max())
					####remove blanks
					sheets[j].iloc[:,i]=sheets[j].iloc[:,i].apply(lambda x: x.strip() if type(x)==str else x)
				##count max number of characteres per field
				chars=pd.concat([pd.Series(sheets[j].columns).rename('FIELD'),pd.Series(count_chars).rename('CHARACTERES')],axis=1)
				common_fields_chars=common_fields[j].merge(chars,on='FIELD',how='outer')
				c=common_fields_chars.iloc[:,[0,1,3,2]]
				c=c[c['CHARACTERES']>c['ALLOWED']]
				#c.rename(columns={'FIELD':''},inplace=True)
				if len(c)>0:
					print('','Fields exceeding the number of characteres: '.upper(),'-'*len('Fields exceeding the number of characteres:'),tabulate(c,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
	
				else:
					None
				###all blanks
				#blanks=pd.concat([pd.Series(sheets[j].columns).rename('FIELD'),pd.Series(blank_cases).apply(lambda x: len(x)).rename('COUNT'),pd.Series(blank_cases).rename('Case example'),],axis=1)
				#if np.shape(blanks[blanks['Case example']!='None'])[0]>0:
				#	blank_spaces=blanks[blanks['Case example']!='None']
				#	blank_spaces=blank_spaces.drop(columns=['Case example'])
				blanks=pd.concat(blank_cases)
				blanks=blanks[blanks['COUNT']>0]
				if len(blanks)>0:
					#print('','Fields with blanks spaces: (Blanks Auto Removed) '.upper(),'-'*len('Fields with blanks spaces: (Blanks Auto Removed) '),tabulate(blanks,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n')

					print('','Fields with blanks spaces: (Blanks Auto Removed) '.upper(),'-'*len('Fields with blanks spaces: (Blanks Auto Removed) '),tabulate(blanks,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				###all nulls###
				if np.shape(null_columns.isnull().sum())[0]>0:
					null_fields=pd.DataFrame(null_columns.isnull().sum())
					null_fields.reset_index(level=0, inplace=True)
					null_fields.rename(columns={'index':'FIELD',0:'COUNT'},inplace=True)
					null_fields=null_fields.merge(all_fields.iloc[:,:-1],on='FIELD',how='inner').drop_duplicates()
					#null_fields.rename(columns={'FIELD':''},inplace=True)
					if (null_fields['MANDATORY']=='No').any():
						null_fields_not_mand=null_fields[null_fields['MANDATORY']=='No']
						print('','Fields with Null Values: '.upper(),'-'*len('Fields with Null Values:'),tabulate(null_fields_not_mand,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None
					if (null_fields['MANDATORY']=='Yes').any():
						null_fields_mand=null_fields[null_fields['MANDATORY']=='Yes']
						print('','Fields with Null Values: '.upper(),'-'*len('Fields with Null Values:'),tabulate(null_fields_mand,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None
				else:
					None
				###duplicated rows
				if np.shape(sheets[j][sheets[j].duplicated()])[0]>0:
					dup_rows=np.shape(sheets[j][sheets[j].duplicated()])[0]
					print('','Number of Duplicated rows (Auto Removed): '.upper()+str(dup_rows),'-'*len('Number of Duplicated rows (Auto Removed):'),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				###update company Name
				sheets[j]['Company']=company
				##remove duplicated rows
				sheets[j].drop_duplicates(inplace=True)
				#replace NA by empty
				sheets[j].fillna('',inplace=True)
				####locations
				filtered_locations=sheets[j].filter(regex=re.compile('REG|GROUP',re.IGNORECASE))
				length=[]
				loc_issues=[]
				#print('','Count of distinct Locations:'.upper(),'-'*len('Count of distinct Locations:'),sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
				for k in range(np.shape(filtered_locations)[1]):
					length.append(len(filtered_locations.iloc[:,k].unique()))
					summary_list.append(pd.concat([pd.Series('COUNT OF DISTINCT '+filtered_locations.columns[k]).str.upper(),pd.Series(length[k])],axis=1))
					#print(filtered_locations.columns[k]+': ' + str(length[k]) ,'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
					if filtered_locations.iloc[:,k].str.isupper().any():
						loc_issues.append(pd.concat([pd.Series(filtered_locations.columns[k]),pd.Series('upper and lower case for same location')],axis=1))
						#print('','Location values:'.upper(),'-'*len('Location values:'),'Found Upper cases in ' + filtered_locations.columns[k],'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None
				print('',tabulate(pd.concat(summary_list),tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
				
				###########################################################################################################################
				#SITES
				if (~sheets[j].columns.str.contains('CI N',case=False).any()):
					sites=sheets[j]
					##status
					options=['Proposed','Enabled','Offline','Obsolete','Archive','Delete']
					if len(list(set(sites['Status']) - set(options)))>0:
						sites['Status'][sites['Status']=='']="NULL"

						issues=sites['Status'][sites['Status'].isin(list(set(sites['Status']) - set(options)))]
						issues.rename('issues',inplace=True)
						sites['Status'][sites['Status'].isin(list(set(sites['Status']) - set(options)))]='Enabled'
						fixed_status=pd.concat([issues,sites['Status']],axis=1).drop_duplicates().dropna()
						fixed_status.rename(columns={fixed_status.columns[0]:'OLD STATUS','Status': 'UPDATED STATUS'},inplace=True)
	
						print('','WRONG STATUS VALUES: (AUTO FIXED)','-'*len('WRONG STATUS VALUES: (AUTO FIXED)'),tabulate(fixed_status,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))
					else:
						None
					##special characteres
					sites_chars_changes=pd.DataFrame()
					sites_chars=sites['Site Name'][sites['Site Name'].astype(str).str.contains("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\",regex=True)]
					if np.shape(sites_chars)[0]>0:
						sites_chars_corrected=sites_chars.astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))
						sites_chars_changes=pd.concat([sites_chars,sites_chars_corrected.rename('Possible corrected Site Name')],axis=1)
						#sites['Site Name']=sites['Site Name'].astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))					#cis['CI Name']=cis['CI Name'].astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))
						print('','Sites with Special Characteres (Auto fixing to be implemented): '.upper()+ str(len(sites_chars)),'-'*len('CIs with Special Characteres (Auto fixed):'),tabulate(sites_chars_changes.head(),headers='keys',tablefmt='fancy_grid',showindex=False),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))

					else:
						None
					###check coordinates
					sites['Longitude_correct']=sites['Longitude'].astype(str).apply(lambda x: re.findall(r'-?\d+\.?\d*', re.sub('[^A-Za-z0-9]+', '.',x))).apply(lambda x: ''.join(x) if len(x)>0 else 'NaN').astype(float)
					sites['Latitude_correct']= sites['Latitude'].astype(str).apply(lambda x: re.findall(r'-?\d+\.?\d*', re.sub('[^A-Za-z0-9]+', '.',x))).apply(lambda x: ''.join(x) if len(x)>0 else 'NaN').astype(float)
					if not (sites['Latitude_correct'].equals(sites['Latitude']) or sites['Longitude_correct'].equals(sites['Longitude'])):
						sites['Latitude']= sites['Latitude_correct']
						sites['Longitude']=sites['Longitude_correct']
						loc_issues.append(pd.concat([pd.Series('COORDINATES (Latitude,Longitude)'),pd.Series('Decimal delimiter should be a commma (Auto fixed)')],axis=1))
						#print('','COORDINATES (Latitude,LONGITUDE): '+'Decimal delimiter should be a commma (Auto fixed).','-'*len('COORDINATES (Latitude,LONGITUDE):'),'',sep='\n',file=open(report +'warningsSites.txt','a',encoding='utf8'))
					else:
						None
					if len(loc_issues)>0:
						print('','LOCATIONS ISSUES:','-'*len('LOCATIONS ISSUES:'),tabulate(pd.concat(loc_issues),headers=['LOCATION','ISSUE'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))

					else:
						None

					#print(tabulate(pd.concat(loc_issues),headers=['LOCATION','ISSUE'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))
					sites.drop(columns=['Latitude_correct','Longitude_correct'],axis=1,inplace=True)
					###compare sites to upload with existing sites
					existing_sites_list=[]
					wrong_locations_sites_list=[]
					new_sites=pd.DataFrame()
					count_issues_sites=[]
					####duplicated sites
					sites_no_alias=sites.drop(sites.filter(regex=re.compile('Alias',re.IGNORECASE)).columns.tolist(),axis=1)
					sites_no_alias.drop_duplicates(inplace=True)
					sites_locations=sites_no_alias.filter(regex=re.compile('SITE N|REG|GROUP',re.IGNORECASE))
					dup_sites=pd.DataFrame(sites_no_alias[pd.DataFrame(sites_locations.iloc[:,0]).duplicated(keep=False)].drop_duplicates()).sort_values([sites_locations.columns[0]])
					duplicate_sites=[]
					duplicate_sites.append(dup_sites.groupby(pd.DataFrame(sites_locations.iloc[:,0]).columns.values[0]).size().reset_index(name='counts'))
					if len(duplicate_sites[0])>0:
						count_issues_sites.append(pd.concat([pd.Series('Duplicated Site Names (excluding duplicate rows):'.upper()),pd.Series(len(duplicate_sites[0]))],axis=1))
					else:
						None
					#lookup sites in ITSM
					if np.shape(sites_itsm)[0]>0:

						all_sites=sites.merge(sites_itsm.set_index(sites_itsm.columns[0]),left_on=sites.filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='outer',indicator=True).drop_duplicates()
						###existing_sites##
						existing_sites=all_sites[all_sites['_merge']=='both'].iloc[:,:-1]
						new_sites=all_sites[all_sites['_merge']=='left_only'].iloc[:,:-1]
						###pick Site
						existing_sites=pd.concat([existing_sites.filter(regex=re.compile('SITE',re.IGNORECASE)).iloc[:,0],existing_sites.filter(regex=re.compile('REG|SITE GROUP|CITY',re.IGNORECASE))],axis=1)
						existing_sites.rename(columns={
							existing_sites.filter(regex=re.compile('REGION',re.IGNORECASE)).columns[0]:'Region in Sites',
							existing_sites.filter(regex=re.compile('REGION',re.IGNORECASE)).columns[1]:'Region in ITSM Sites',
							existing_sites.filter(regex=re.compile('SITE GROUP',re.IGNORECASE)).columns[0]:'Site Group in Sites',
							existing_sites.filter(regex=re.compile('SITE GROUP',re.IGNORECASE)).columns[1]:'Site Group in ITSM Sites'},
							inplace=True)
						
						if np.shape(existing_sites)[0]>0:
							###new sites
							#print('','New sites to upload in CMDB: '.upper() + str(np.shape(new_sites)[0]),'-'*len('New Sites to upload in CMDB:'),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
							wrong_locations_sites=pd.concat([existing_sites[(existing_sites['Region in Sites'])!=(existing_sites['Region in ITSM Sites'])],existing_sites[(existing_sites['Site Group in Sites'])!=(existing_sites['Site Group in ITSM Sites'])]],axis=0).drop_duplicates()
							correct_locations_sites=pd.concat([existing_sites[(existing_sites['Region in Sites'])==(existing_sites['Region in ITSM Sites'])],existing_sites[(existing_sites['Site Group in Sites'])==(existing_sites['Site Group in ITSM Sites'])]],axis=0).drop_duplicates()
							###already existing sites

							if len(correct_locations_sites)>0:
								count_issues_sites.append(pd.concat([pd.Series('Existing sites with same locations:'.upper()),pd.Series(len(correct_locations_sites))],axis=1))
							
								#print('','Existing sites with same locations: '.upper() +str(np.shape(correct_locations_sites)[0]),'-'*len('Existing sites with same locations:'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
								existing_sites_list.append(correct_locations_sites)
							else:
								None
							###existing sites with different locations
							if len(wrong_locations_sites)>0:
								count_issues_sites.append(pd.concat([pd.Series('Existing sites with mismatched locations:'.upper()),pd.Series(len(wrong_locations_sites))],axis=1))

								#rint('','Existing sites with mismatched locations: '.upper() +str(np.shape(wrong_locations_sites)[0]),'-'*len('Existing sites with mismateched locations:'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
								wrong_locations_sites_list.append(wrong_locations_sites)
							else:
								None
							

						else:
							None
					else:
						None				
					
					if len(count_issues_sites)>0:
						print('','DUPLICATION AND LOCATIONS ISSUES:','-'*len('DUPLICATION AND LOCATIONS ISSUES:'),tabulate(pd.concat(count_issues_sites),headers=['ISSUE','COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None
					sites.to_csv(report + company + '_SITES_SCREENED_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)
				else:
					None
				###CIS
				if (sheets[j].columns.str.contains('CI N',case=False).any()):
					cis=sheets[j]
					##status
					options=["Ordered","Received","Being Assembled","Deployed","In Repair","Down","End of Life","Transferred","Delete","In Inventory","On Loan","Disposed","Reserved","Return to Vendor"]
					if len(list(set(cis['Status']) - set(options)))>0:
						cis['Status'][cis['Status']=='']="NULL"
						issues=cis['Status'][cis['Status'].isin(list(set(cis['Status']) - set(options)))]
						issues.rename('issues',inplace=True)
						cis['Status'][cis['Status'].isin(list(set(cis['Status']) - set(options)))]='Deployed'
						fixed_status=pd.concat([issues,cis['Status']],axis=1).drop_duplicates().dropna()
						fixed_status.rename(columns={fixed_status.columns[0]:'OLD STATUS','Status': 'UPDATED STATUS'},inplace=True)
	
						print('','WRONG STATUS VALUES: (AUTO FIXED)','-'*len('WRONG STATUS VALUES: (AUTO FIXED)'),tabulate(fixed_status,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))
					else:
						None
					#priority
					#priority
					cis['Priority']=cis['Priority'].str.upper()
					cis['Priority'][cis['Priority']=='']="NULL"
					prios=['PRIORITY_'+str(x) for x in list(range(1,6))]
					if len(list(set(cis['Priority']) - set(prios)))>0:
						issues=cis['Priority'][cis['Priority'].isin(list(set(cis['Priority']) - set(prios)))]
						issues.rename(columns={'Priority':'issues'},inplace=True)
						cis['Priority'][cis['Priority'].isin(list(set(cis['Priority']) - set(prios)))]='PRIORITY_5'
						fixed_prios=pd.concat([issues,cis['Priority']],axis=1).drop_duplicates().dropna()
						fixed_prios.rename(columns={fixed_prios.columns[0]:'OLD PRIORITY','Priority': 'UPDATED PRIORITY'},inplace=True)
	
						print('','WRONG PRIORITY VALUES: (AUTO FIXED)','-'*len('WRONG PRIORITY VALUES: (AUTO FIXED)'),tabulate(fixed_prios,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))
					else:
						None
					cis_chars_changes=pd.DataFrame()
					cis_chars=cis['CI Name'][cis['CI Name'].astype(str).str.contains("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\",regex=True)]
					if np.shape(cis_chars)[0]>0:
						cis_chars_corrected=cis_chars.astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))
						cis_chars_changes=pd.concat([cis_chars,cis_chars_corrected.rename('Possible correct CI Name')],axis=1)
						#cis['CI Name']=cis['CI Name'].astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))					#cis['CI Name']=cis['CI Name'].astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))
						print('','CIs with Special Characteres (Auto fixing to be implemented): '.upper()+ str(len(cis_chars)),'-'*len('CIs with Special Characteres (Auto fixed):'),tabulate(cis_chars_changes.head(),headers='keys',tablefmt='fancy_grid',showindex=False),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))

					else:
						None
					###existing cis
					#errors=pd.DataFrame(columns=['COUNT'])
					existing_cis=pd.DataFrame()
					if len(cis_itsm)>0:
						existing_cis=cis.merge(cis_itsm,on='CI Name',how='inner')
						if len(existing_cis)>0:
							ex_cis=pd.Series(len(existing_cis))
							count_issues.append(pd.concat([pd.Series('CIS ALREADY EXISTING IN CMDB:'.upper()),ex_cis],axis=1))
					#	#print('','NUMBER OF CIS ALREADY EXISTING IN CMDB:','-'*len('NUMBER OF CIS ALREADY EXISTING IN CMDB:'),tabulate(ex_cis,headers=['COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
					#	#print('NUMBER OF CIS ALREADY EXISTING IN CMDB: '+str(len(existing_cis)), '-'*len('NUMBER OF CIS ALREADY EXISTING IN CMDB:'),
					#
						else:
							None
					else:
						None
					##DUPLICATED CIS
					filtered_cis=cis.filter(regex=re.compile('CI N',re.IGNORECASE))
					duplicate_cis=[]
					###change####
					dup_cis=pd.DataFrame(cis[filtered_cis.duplicated(keep=False)].drop_duplicates()).sort_values([filtered_cis.columns[0]])
					duplicate_cis.append(dup_cis.groupby(filtered_cis.columns.values[0]).size().reset_index(name='counts'))
					if len(dup_cis)>0:
						dup_cis_count=pd.Series(len(duplicate_cis[0]))
						count_issues.append(pd.concat([pd.Series('Duplicated CI Names (excluding duplicate rows):'.upper()),dup_cis_count],axis=1))
						#print('','Duplicated CI Names (excluding duplicate rows): '.upper(),'-'*len('Duplicated CI Names (excluding duplicate rows): '.upper()),tabulate(dup_cis_count,headers=['COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))

						#print('','Duplicated CI Names (excluding duplicate rows): '.upper(),'-'*len('Duplicated CI Names (excluding duplicate rows):'),tabulate(dup_sites_cis.set_index(dup_sites_cis.columns[0]),headers=['COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None


					###duplicated region per site in CIs
					sites_reg=cis[['Site','Region','Site Group']].drop_duplicates()
					dup_sites_reg=sites_reg[sites_reg.iloc[:,0].duplicated(keep=False)].sort_values([sites_reg.columns[0]])
					cis_locations=[]
					if np.shape(dup_sites_reg)[0]>0:
						cis_locations.append(dup_sites_reg)
						dup_sites_cis=pd.Series(len(dup_sites_reg))
						count_issues.append(pd.concat([pd.Series('Duplicated Location per Site in CIs data:'.upper()),dup_sites_cis],axis=1))
						#print('','Duplicated Location per Site in CIs data: '.upper(),'-'*len('Duplicated Location per Site in CIs data:'),tabulate(dup_sites_cis,headers=['COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None

					

					##check DNS, CI Description and Domain
					filtered_desc=cis.filter(regex=re.compile('DESC',re.IGNORECASE))
					filtered_region=cis.filter(regex=re.compile('REG',re.IGNORECASE))
					filtered_site_group=cis.filter(regex=re.compile('SITE GROUP',re.IGNORECASE))
					filtered_dns=cis.filter(regex=re.compile('DNS',re.IGNORECASE))      
					with_region=(cis[filtered_desc.columns[0]]==cis[filtered_cis.columns[0]] + ' | ' + cis[filtered_region.columns[0]]).all()
					with_sitegroup=(cis[filtered_desc.columns[0]]==cis[filtered_cis.columns[0]] + ' | ' + cis[filtered_site_group.columns[0]].astype(str)).all()
					well_conc=[with_region,with_sitegroup]
					conc_fields=[]
					if well_conc[0]==False & well_conc[1]==False:
						cis['Suggested '+filtered_desc.columns[0]]=cis[filtered_cis.columns[0]] + ' | ' + cis[filtered_region.columns[0]]
						conc_fields.append(pd.concat([pd.Series('CI Description'.upper()),pd.Series('CI Name | Location/System Role/Product Name')],axis=1))
						#print('','CI Description concatenation:'.upper(),'-'*len('CI Description concatenation:'),'A concatenation with Region is suggested. Please check the Data Model','',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
					else:
						None
					
					well_conc_dns=(cis[filtered_dns.columns[0]]==cis[filtered_cis.columns[0]]).all()
					if well_conc_dns==False:
						cis['Suggested '+ filtered_dns.columns[0]]=cis[filtered_cis.columns[0]]
						conc_fields.append(pd.concat([pd.Series('DNS Host Name values'.upper()),pd.Series('Replaced by CI Name values')],axis=1))

						#print('','DNS Host Name values:'.upper(),'-'*len('DNS Host Name values:'),'Replaced by CI Name values','',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))       
			
					else:
						None
					
					if len(conc_fields)>0:
						print('','CONCATENATED FIELDS','-'*len('CONCATENATED FIELDS'),tabulate(pd.concat(conc_fields),headers=['FIELD','SUGGESTION'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))
					else:
						None
					###check product catalogue
					opcat_template='Prod_Cats_V2'
					
					template=pd.read_csv(glob.glob(opcat_template+'/*.csv')[0],sep=';')
					template.drop(columns=['Model Version'],axis=1,inplace=True)
					template=template.drop_duplicates()
					template.rename(columns={template.columns[3]:'Product Name'},inplace=True)
					template.rename(columns={template.columns[0]:'Tier 1 in catalogue',
						template.columns[1]:'Tier 2 in catalogue',
						template.columns[2]:'Tier 3 in catalogue',
									template.columns[4]:'Manufacturer in catalogue'},
								 inplace=True)

					prodcats_cis=cis[['Tier 1','Tier 2','Tier 3','Product Name','Manufacturer']]
					prod_missing=prodcats_cis.loc[~prodcats_cis['Product Name'].isin(template.iloc[:,3])].drop_duplicates()
					prod_name=prod_missing.filter(regex=re.compile('Product N',re.IGNORECASE)).iloc[:,0]
					prod_match=prod_name.apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else get_close_matches(x,template.iloc[:,3].astype(str).unique().tolist()))
					prod_suggested=prod_name.apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else get_close_matches(x,template.iloc[:,3].astype(str).unique().tolist(),1)).apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else ''.join(x))
					prod_missing_final=pd.concat([prod_missing,prod_suggested.rename('Suggested Product Name'),prod_match.rename('Others PN match')],axis=1)
					same_prod=prodcats_cis[~prodcats_cis['Product Name'].isna()].merge(template,left_on=prod_name.name,right_on='Product Name',how='inner').drop_duplicates()
					same_prod=same_prod.iloc[:,[0,5,1,6,2,7,3,4,8]]
					wrongcats1=same_prod[(same_prod.iloc[:,0]!=same_prod.iloc[:,1]) & (same_prod.iloc[:,2]==same_prod.iloc[:,3]) & (same_prod.iloc[:,4]==same_prod.iloc[:,5])].drop_duplicates()
					wrongcats1['FIELD'] ='Tier 1'
					wrongcats2=same_prod[(same_prod.iloc[:,2]!=same_prod.iloc[:,3]) & (same_prod.iloc[:,0]==same_prod.iloc[:,1]) & (same_prod.iloc[:,4]==same_prod.iloc[:,5])].drop_duplicates()
					wrongcats2['FIELD'] ='Tier 2'
					wrongcats3=same_prod[(same_prod.iloc[:,4]!=same_prod.iloc[:,5]) &(same_prod.iloc[:,0]==same_prod.iloc[:,1]) & (same_prod.iloc[:,2]==same_prod.iloc[:,3])].drop_duplicates()
					wrongcats3['FIELD'] ='Tier 3'
					wrongcats1and2=same_prod[(same_prod.iloc[:,0]!=same_prod.iloc[:,1]) & (same_prod.iloc[:,2]!=same_prod.iloc[:,3]) & (same_prod.iloc[:,4]==same_prod.iloc[:,5])].drop_duplicates()
					wrongcats1and2['FIELD'] ='Tier 1 and 2'
					wrongcats1and3=same_prod[(same_prod.iloc[:,0]!=same_prod.iloc[:,1]) & (same_prod.iloc[:,2]==same_prod.iloc[:,3]) & (same_prod.iloc[:,4]!=same_prod.iloc[:,5])].drop_duplicates()
					wrongcats1and3['FIELD'] ='Tier 1 and 3'
					wrongcats2and3=same_prod[(same_prod.iloc[:,2]!=same_prod.iloc[:,3]) & (same_prod.iloc[:,4]!=same_prod.iloc[:,5]) & (same_prod.iloc[:,0]==same_prod.iloc[:,1])].drop_duplicates()
					wrongcats2and3['FIELD'] ='Tier 2 and 3'
					wrongcats1and2and3=same_prod[(same_prod.iloc[:,0]!=same_prod.iloc[:,1]) & (same_prod.iloc[:,2]!=same_prod.iloc[:,3]) & (same_prod.iloc[:,4]!=same_prod.iloc[:,5])].drop_duplicates()
					wrongcats1and2and3['FIELD'] ='Tier 1, Tier 2 and 3'
					wrong_manufacturer=same_prod[(same_prod.iloc[:,7]!=same_prod.iloc[:,8])].drop_duplicates()    
					wrong_manufacturer['FIELD'] ='Manufacturer'
					wrong_Tiers=pd.concat([wrongcats1,wrongcats2,wrongcats3,wrongcats1and2,wrongcats1and3,wrongcats2and3,wrongcats1and2and3,wrong_manufacturer],axis=0)
					wrong_t=pd.DataFrame(wrong_Tiers["FIELD"])
					wrong_t=wrong_t.groupby(wrong_t.columns.values[0]).size().reset_index(name='counts')
					wrong_prod=pd.concat([pd.Series('Product Name').rename('FIELD'),pd.Series(len(prod_missing_final['Product Name'])).rename('COUNT')],axis=1)
					#wrong_prod=pd.concat([pd.Series('Product Name').rename('FIELD'),pd.Series(len(prod_missing_final['Product Name'])).rename('Count')],axis=1)
					wrong_t=pd.DataFrame(wrong_Tiers["FIELD"])
					wrong_t=wrong_t.groupby(wrong_t.columns.values[0]).size().reset_index(name='COUNT')
					wrong_catalogue=pd.concat([wrong_prod,wrong_t],axis=0)
					wrong_catalogue.reset_index(inplace = True,drop =True)
					wrong_catalogue=wrong_catalogue[wrong_catalogue['COUNT']>0]
					#wrong_catalogue.rename(columns={'FIELD':''},inplace=True)
					#wrong_catalogue=wrong_catalogue.set_index(wrong_catalogue.columns[0])
					if len(wrong_catalogue)>0:
						print('','WRONG VALUES IN PRODUCT CATEGORIZATION:','-'*len('WRONG VALUES IN PRODUCT CATEGORIZATION:'), tabulate(wrong_catalogue,headers="keys",tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))              
					else:
						None
					if len(loc_issues)>0:
						print('','LOCATIONS ISSUES:','-'*len('LOCATIONS ISSUES:'),tabulate(pd.concat(loc_issues),headers=['LOCATION','ISSUE'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'warnings' + names[j] + '.txt','a',encoding='utf8'))
					else:
						None
					cis.to_csv(report + company + '_CIS_SCREENED_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")+'.csv',sep=';',mode='w',index=False)

				else:
					None
			
			if (len(cis)>0):			
				cis_sites_locations=[]
				if (len(sites)>0):
					conc_sites=[]
					for i in range(len(sites_itsm.columns)):
						conc_sites.append(pd.concat([sites_itsm.iloc[:,i],sites_locations.iloc[:,i]],axis=0))
					all_sites2=pd.concat(conc_sites,axis=1)
					all_sites2.rename(columns={
					all_sites2.columns[0]:'Site Name',
					all_sites2.columns[1]:'Region',
					all_sites2.columns[2]:'Site Group'},inplace=True)
				else:
					all_sites2=sites_itsm
				##CIs with non existing sites
				new_sites_in_cis=sites_reg.merge(all_sites2.filter(regex=re.compile('SITE N|SITE+|SITE*|REG|GROUP',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=sites_reg.filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='outer',indicator=True).drop_duplicates()
				new_sites_in_cis=new_sites_in_cis[new_sites_in_cis['_merge']=='left_only'].iloc[:,:-1]
				if np.shape(new_sites_in_cis)[0]>0:
					new_sites_in_cis.rename(columns={new_sites_in_cis.columns[1]:'Region',
									 new_sites_in_cis.columns[2]:'Site Group'},
								inplace=True
								   )
					new_sites_in_cis=new_sites_in_cis[['Site','Region','Site Group']]
					count_issues.append(pd.concat([pd.Series('CIs with non existing sites:'.upper()),pd.Series(len(new_sites_in_cis))],axis=1))

					#print('','CIs with non existing sites: '.upper()+str(np.shape(new_sites_in_cis)[0]),'-'*len('CIs with non existing sites:'),'',sep='\n',file=open(report +'errorsCIs.txt','a',encoding='utf8'))
				else:
					None
				existing_sites2=sites_reg.merge(all_sites2.filter(regex=re.compile('SITE N|REG|SITE GROUP|CITY',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=sites_reg.filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='inner').drop_duplicates()
				existing_sites2.rename(columns={
					existing_sites2.filter(regex=re.compile('REGION',re.IGNORECASE)).columns[0]:'Region in CIs',
					existing_sites2.filter(regex=re.compile('REGION',re.IGNORECASE)).columns[1]:'Region in Sites',
					existing_sites2.filter(regex=re.compile('SITE GROUP',re.IGNORECASE)).columns[0]:'Site Group in CIs',
					existing_sites2.filter(regex=re.compile('SITE GROUP',re.IGNORECASE)).columns[1]:'Site Group in Sites'},inplace=True)
				wrong_locations=pd.concat([existing_sites2[(existing_sites2['Region in CIs'])!=(existing_sites2['Region in Sites'])],existing_sites2[(existing_sites2['Site Group in CIs'])!=(existing_sites2['Site Group in Sites'])]],axis=0).drop_duplicates()
				if np.shape(wrong_locations)[0]>0:
					cis_sites_locations.append(wrong_locations)
					count_issues.append(pd.concat([pd.Series('Mismatched Locations between CIs and Sites data:'.upper()),pd.Series(len(wrong_locations))],axis=1))

					#print('','Mismatched Locations between CIs and Sites data: '.upper()+ str(np.shape(wrong_locations)[0]),'-'*len('Mismatched Locations between CIs and Sites data:'),'',sep='\n',file=open(report +'errorsCIs.txt','a',encoding='utf8'))
				else:
					None
				
			else:
				None
			if len(count_issues)>0:				
				print('','DUPLICATION AND LOCATIONS ISSUES:','-'*len('DUPLICATION AND LOCATIONS ISSUES:'),tabulate(pd.concat(count_issues),headers=['ISSUE','COUNT'],tablefmt="fancy_grid",showindex=False),'',sep='\n',file=open(report +'errorsCIs.txt','a',encoding='utf8'))
			else:
				None


		with pd.ExcelWriter(report + company + '_ERRORS_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
			if len(sites)>0:
				#sites.to_excel(writer, 'sites',index=False)
				if np.shape(new_sites)[0]>0:
					new_sites.to_excel(writer,'New Sites',index=False)
				else:
					None    		
				if np.shape(dup_sites)[0]>0:
					dup_sites.to_excel(writer, '1.Duplicate Sites',index=False)
				else:
					None
				if len(existing_sites_list)>0:
					existing_sites_list[0].to_excel(writer, '2.Existing Sites',index=False)
				else:
					None
				if len(wrong_locations_sites_list)>0:
					wrong_locations_sites_list[0].to_excel(writer, '3.Location Issues in Sites',index=False)
				else:
					None
				if np.shape(sites_chars_changes)[0]>0:
					sites_chars_changes.to_excel(writer, '4.Special Characteres in Sites',index=False)
				else:
					None
			else:
				print('','No Sites data to validate'.upper(),sep='\n',file=open(report +'summary' + names[j] + 'Sites.txt','a',encoding='utf8'))
			if len(cis)>0:
				#cis.to_excel(writer, 'cis',index=False)
				if len(existing_cis)>0:
					existing_cis.to_excel(writer,'5.CIs already in CMDB',index=False)
				else:
					None

				if np.shape(dup_cis)[0]>0:
					dup_cis.to_excel(writer, '6.Duplicate CIs',index=False)
				else:
					None

				if np.shape(new_sites_in_cis)[0]>0:
					new_sites_in_cis.to_excel(writer,'7.CIs with non existing sites',index=False)
				else:
					None
				if len(cis_locations)>0:
					cis_locations[0].to_excel(writer, '8.Duplicated Location in CIs',index=False)
				else:
					None
				
				if len(cis_sites_locations)>0:
					cis_sites_locations[0].to_excel(writer, '9.Location Issues in CIS',index=False)
				else:
					None
				
				
				
				if np.shape(cis_chars_changes)[0]>0:
					cis_chars_changes.to_excel(writer, '10.Special Characteres in CIs',index=False)
				else:
					None
					
				if len(wrong_Tiers)>0:
					wrong_Tiers.to_excel(writer,'11.wrong_tiers',index=False)
				else:
					None
				if len(prod_missing_final)>0:
					prod_missing_final.to_excel(writer,'12.wrong_product_name',index=False)
	
				else:
					None
				
				  
			else:
				print('','No CIs data to validate'.upper(),sep='\n',file=open(report +'summary' + names[j] + 'CIs.txt','a',encoding='utf8'))
			writer.save()	
