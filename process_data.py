import numpy as np
import pandas as pd
import re
import os
import glob
from difflib import get_close_matches
import itertools
#import sys
import datetime as dt
#from io import StringIO
#import csv

def process_file(path,company,report):
	startTime = dt.datetime.now()
	sites_itsm=pd.DataFrame(columns=['Site Name','Region','Site Group'])
	cis_itsm=pd.DataFrame(columns=['CI Name'])
	all_sites=pd.DataFrame(columns=['Site','Region','Site Group'])

	######read files
	sheets=[]
	cis=pd.DataFrame()
	sites=pd.DataFrame()
	files=pd.DataFrame()
	for j in range(len(glob.glob(path+'/*'))):
		if glob.glob(path+'/*')[j].endswith(('.xls','.xlsx')):
			files=pd.read_excel(glob.glob(path+'*')[j],sheet_name=None)
			for frame in files.keys():
				sheets.append(files[frame])
		elif glob.glob(path+'/*')[j].endswith('.csv'):
			sheets.append(pd.read_csv(glob.glob(path+'*')[j],sep=";",encoding='ISO-8859-1'))
		else:
			None
	
	##check correct fields
	sites_fields=['Company','Site Name','Site Alias','Description','Region','Site Group','Street','Country','City','Latitude','Longitude','Location ID','Additional Site Details','Maintenance Circle Name','Site Type','Status']
	cis_fields=['Company','CI Name','CI Description','Tag Number','System Role','Status','Priority','Additional Information','Tier 1','Tier 2','Tier 3','Product Name','Model Version','Manufacturer','Region','Site Group','Site','DNS Host Name','Domain','CI ID']
	fields=sites_fields+cis_fields
	char_num=[254,60,60,255,60,60,90,60,60,12,12,30,10,70,22,8,254,254,254,64,30,16,10,254,60,60,60,254,254,254,60,60,60,254,254,64]
	field_type=[['Yes']*3,['No'],['Yes']*2,['No'],['Yes']*2,['No']*6,['Yes']*3,['No']*3,['Yes']*2,['No'],['Yes']*4,['No'],['Yes']*4,['No']*3]
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
		unmatched=pd.concat([pd.Series(['']*len(pd.Series(list(itertools.chain(*unmatched_fields))))).rename(''),pd.Series(list(itertools.chain(*unmatched_fields))).rename('FIELD')],axis=1)
		print('',unmatched.set_index('FIELD'),'',sep='\n',file=open(report +'Mismatched_fields.txt','a',encoding='utf8'))
		#print('','No Warnings to display'.upper(),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
	else:
		####cmdb inventory
		if pd.Series(os.listdir('CMDB_inventory')).str.contains(company).any():
			sites_itsm=pd.read_csv(glob.glob('CMDB_inventory/'+company+'*Sites_report.csv')[0],sep=';',encoding='ISO-8859-1')
			sites_itsm=sites_itsm[['Site Name', 'Region', 'Site Group']]
			cis_itsm=pd.read_csv(glob.glob('CMDB_inventory/'+company+'*CIs_report.csv')[0],sep=';',encoding='ISO-8859-1')
			cis_itsm=cis_itsm[['CI Name']]
			counts_cmb=pd.concat([pd.Series(['Sites','CIs']).rename(''),pd.Series([len(sites_itsm),len(cis_itsm)]).rename('COUNT')],axis=1).set_index('')
			print('',counts_cmb,'',sep='\n',file=open(report +'summary_CMDB.txt','a',encoding='utf8'))      
		else:
			print('bla')
			None#print('Missing CMDB inventory','',sep='\n',file=open(report +'summary_CMDB.txt','a',encoding='utf8'))
		##############
		all_fields=pd.concat([pd.Series(fields).rename('FIELD'),pd.Series(field_type).rename('MANDATORY'),pd.Series(char_num).rename('ALLOWED')],axis=1)
		common_fields=[]
		for j in range(len(sheets)):
			###first overview prints####
			print(sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
			print(sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
			print(sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
			if len(os.listdir(path))==1:
				print('','INPUT FILE NAME: '+os.listdir(path)[0],'-'*len('INPUT FILE NAME:'),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
			else:
				print('','INPUT FILE NAME: '+os.listdir(path)[j],'-'*len('INPUT FILE NAME:'),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))

			#ITSM SITES report
			#itsm_columns=['Site Name','Region','Site Group','City']
			#if len(glob.glob(history+'/*'))==0:
			#	sites_itsm=pd.DataFrame(columns=itsm_columns)
			#	sites_itsm=sites_itsm[itsm_columns]
			#	print('','All Existing Sites in CMDB: '.upper(),'-'*len('All Existing Sites in CMDB: '),'No CMDB Report: Suggest to get a full report of existing sites','',sep='\n',file=open(report +'warnings'+ names[j] + '.txt','a',encoding='utf8'))
			#else:
			#	sites_itsm=pd.read_csv(glob.glob(history+'/*Sites_report.csv')[0],sep=';')
			#	#sites_itsm=pd.read_excel(glob.glob(history+'/*')[0],pd.ExcelFile(glob.glob(history+'/*')[0]).sheet_names[0])
			#	cis_itsm=pd.read_csv(glob.glob(history+'/*CIs_report.csv')[0],sep=';')
			#	cis_itsm=pd.DataFrame(cis_itsm['CI Name'])
			#	#cis_itsm=pd.read_excel(glob.glob(history+'/*')[0],pd.ExcelFile(glob.glob(history+'/*')[0]).sheet_names[0])
			#	sites_itsm=sites_itsm.filter(regex=re.compile('SITE N|REG|GROUP|CITY',re.IGNORECASE))
			#	sites_itsm.rename(columns={
			#		sites_itsm.columns[0]:'Site Name',
			#		sites_itsm.columns[1]:'Region',
			#		sites_itsm.columns[2]:'Site Group',
			#		sites_itsm.columns[3]:'City'},inplace=True)
			#	print('','All Existing Sites in CMDB: '.upper()+str(np.shape(sites_itsm)[0]),'-'*len('All Existing Sites in CMDB: '),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))      
			
			print('Number of records: '.upper()+str(np.shape(sheets[j])[0]),'-'*len('Number of records:'),'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))					
			common_fields.append(all_fields.merge(pd.DataFrame(pd.Series(sheets[j].columns).rename('FIELD')),on='FIELD',how='inner').drop_duplicates())
			blank_cases=[]
			count_chars=[]
			#check null values
			null_columns=sheets[j][sheets[j].columns[sheets[j].isnull().any()]]  		
			###blanks and char num
			for i in range(len(sheets[j].columns)):
				blank_find=sheets[j].iloc[:,i][sheets[j].iloc[:,i].astype(str).apply(lambda x: x[0].isspace() or x[len(x)-1].isspace())]
				if len(blank_find)>0:
					blank_cases.append(blank_find.iloc[0])

				else:
					blank_cases.append('None')
				####count_chars
				count_chars.append(sheets[j].iloc[:,i].apply(lambda x: x if pd.isnull(x) else (len(str(round(x,5))) if type(x)==float else len(str(x)))).max())
				####remove blanks
				sheets[j].iloc[:,i]=sheets[j].iloc[:,i].apply(lambda x: x.strip() if type(x)==str else x)
			##count max number of characteres per field
			chars=pd.concat([pd.Series(sheets[j].columns).rename('FIELD'),pd.Series(count_chars).rename('CHARACTERES')],axis=1)
			common_fields_chars=common_fields[j].merge(chars,on='FIELD',how='outer')
			c=common_fields_chars.iloc[:,[0,1,3,2]]
			c=c[c['CHARACTERES']>c['ALLOWED']]
			c.rename(columns={'FIELD':''},inplace=True)
			if len(c)>0:
				print('','Fields exceeding the number of characteres: '.upper(),'-'*len('Fields exceeding the number of characteres:'),c.set_index(c.columns[0]),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))

			else:
				None
			###all blanks
			blanks=pd.concat([pd.Series(sheets[j].columns).rename('FIELD'),pd.Series(blank_cases).apply(lambda x: len(x)).rename('COUNT'),pd.Series(blank_cases).rename('Case example'),],axis=1)
			if np.shape(blanks[blanks['Case example']!='None'])[0]>0:
				blank_spaces=blanks[blanks['Case example']!='None']
				blank_spaces=blank_spaces.drop(columns=['Case example'])
				blank_spaces.rename(columns={'FIELD':''},inplace=True)
				print('','Fields with blanks spaces: (Blanks Auto Removed) '.upper(),'-'*len('Fields with blanks spaces: (Blanks Auto Removed)'),blank_spaces.set_index(blank_spaces.columns[0]),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
			else:
				None
			###all nulls###
			if np.shape(null_columns.isnull().sum())[0]>0:
				null_fields=pd.DataFrame(null_columns.isnull().sum())
				null_fields.reset_index(level=0, inplace=True)
				null_fields.rename(columns={'index':'FIELD',0:'COUNT'},inplace=True)
				null_fields=null_fields.merge(all_fields.iloc[:,:-1],on='FIELD',how='inner').drop_duplicates()
				null_fields.rename(columns={'FIELD':''},inplace=True)
				if (null_fields['MANDATORY']=='No').any():
					null_fields_not_mand=null_fields[null_fields['MANDATORY']=='No']
					print('','Fields with Null Values: '.upper(),'-'*len('Fields with Null Values:'),null_fields_not_mand.set_index(null_fields_not_mand.columns[0]),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				if (null_fields['MANDATORY']=='Yes').any():
					null_fields_mand=null_fields[null_fields['MANDATORY']=='Yes']
					print('','Fields with Null Values: '.upper(),'-'*len('Fields with Null Values:'),null_fields_mand.set_index(null_fields_mand.columns[0]),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
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
			##remove duplicated rows
			sheets[j].drop_duplicates(inplace=True)
			#replace NA by empty
			sheets[j].fillna('',inplace=True)
			####locations
			filtered_locations=sheets[j].filter(regex=re.compile('REG|GROUP',re.IGNORECASE))
			length=[]
			print('','Count of distinct Locations:'.upper(),'-'*len('Count of distinct Locations:'),sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
			for k in range(np.shape(filtered_locations)[1]):
				length.append(len(filtered_locations.iloc[:,k].unique()))
				print(filtered_locations.columns[k]+': ' + str(length[k]) ,'',sep='\n',file=open(report +'summary' + names[j] + '.txt','a',encoding='utf8'))
				if filtered_locations.iloc[:,k].str.isupper().any():
					print('','Location values:'.upper(),'-'*len('Location values:'),'Found Upper cases in ' + filtered_locations.columns[k],'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
			###########################################################################################################################
			#SITES
			if (~sheets[j].columns.str.contains('CI N',case=False).any()):
				sites=sheets[j]
				##special characteres
				sites_chars=sites[sites['Site Name'].astype(str).str.contains("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\",regex=True)]
				if np.shape(sites_chars)[0]>0:
					sites['Site Name']=sites['Site Name'].astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))
					print('','Sites with Special Characteres (Auto fixed): '.upper()+str(np.shape(sites_chars)[0]),'-'*len('Sites with Special Characteres (Auto fixed):'),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				###check coordinates
				sites['Longitude_correct']=sites['Longitude'].astype(str).apply(lambda x: re.findall(r'-?\d+\.?\d*', re.sub('[^A-Za-z0-9]+', '.',x))).apply(lambda x: ''.join(x) if len(x)>0 else 'NaN').astype(float)
				sites['Latitude_correct']= sites['Latitude'].astype(str).apply(lambda x: re.findall(r'-?\d+\.?\d*', re.sub('[^A-Za-z0-9]+', '.',x))).apply(lambda x: ''.join(x) if len(x)>0 else 'NaN').astype(float)
				if not (sites['Latitude_correct'].equals(sites['Latitude']) or sites['Longitude_correct'].equals(sites['Longitude'])):
					sites['Latitude']= sites['Latitude_correct']
					sites['Longitude']=sites['Longitude_correct']
					print('','COORDINATES (Latitude,LONGITUDE): '+'Decimal delimiter should be a commma (Auto fixed).','-'*len('COORDINATES (Latitude,LONGITUDE):'),'',sep='\n',file=open(report +'warningsSites.txt','a',encoding='utf8'))
				else:
					None
				sites.drop(columns=['Latitude_correct','Longitude_correct'],axis=1,inplace=True)
				###compare sites to upload with existing sites
				existing_sites_list=[]
				wrong_locations_sites_list=[]
				new_sites=pd.DataFrame()
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
						print('','New sites to upload in CMDB: '.upper() + str(np.shape(new_sites)[0]),'-'*len('New Sites to upload in CMDB:'),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
						wrong_locations_sites=pd.concat([existing_sites[(existing_sites['Region in Sites'])!=(existing_sites['Region in ITSM Sites'])],existing_sites[(existing_sites['Site Group in Sites'])!=(existing_sites['Site Group in ITSM Sites'])]],axis=0).drop_duplicates()
						correct_locations_sites=pd.concat([existing_sites[(existing_sites['Region in Sites'])==(existing_sites['Region in ITSM Sites'])],existing_sites[(existing_sites['Site Group in Sites'])==(existing_sites['Site Group in ITSM Sites'])]],axis=0).drop_duplicates()
						###already existing sites
						if np.shape(correct_locations_sites)[0]>0:
							print('','Existing sites with same locations: '.upper() +str(np.shape(correct_locations_sites)[0]),'-'*len('Existing sites with same locations:'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
							existing_sites_list.append(correct_locations_sites)
						else:
							None
						###existing sites with different locations
						if np.shape(wrong_locations_sites)[0]>0:
							print('','Existing sites with mismatched locations: '.upper() +str(np.shape(wrong_locations_sites)[0]),'-'*len('Existing sites with mismateched locations:'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
							wrong_locations_sites_list.append(wrong_locations_sites)
						else:
							None						
					else:
						None
				else:
					None				
				
				####duplicated sites
				sites_locations=sites.filter(regex=re.compile('SITE N|REG|GROUP',re.IGNORECASE))
				dup_sites=pd.DataFrame(sites[pd.DataFrame(sites_locations.iloc[:,0]).duplicated(keep=False)].drop_duplicates()).sort_values([sites_locations.columns[0]])
				duplicate_sites=[]
				duplicate_sites.append(dup_sites.groupby(pd.DataFrame(sites_locations.iloc[:,0]).columns.values[0]).size().reset_index(name='counts'))
				if np.shape(duplicate_sites[0])[0]>0:
					print('','Duplicated Site Names (excluding duplicate rows): '.upper()+str(len(duplicate_sites[0])),'-'*len('Duplicated Site Names (excluding duplicate rows):'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))  
				else:
					None
			else:
				None
		###CIS
			if (sheets[j].columns.str.contains('CI N',case=False).any()):
				cis=sheets[j]
				###existing cis
				existing_cis=pd.DataFrame()
				existing_cis=cis.merge(cis_itsm,on='CI Name',how='inner')
				if len(existing_cis)>0:
					print('NUMBER OF CIS ALREADY EXISTING IN CMDB: '+str(len(existing_cis)), '-'*len('NUMBER OF CIS ALREADY EXISTING IN CMDB:'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				#cis_chars=pd.DataFrame(cis[cis['CI Name'].astype(str).str.contains("\"|\'|´",regex=True)].drop_duplicates())
				cis_chars=cis[cis['CI Name'].astype(str).str.contains("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\",regex=True)]
				if np.shape(cis_chars)[0]>0:
					cis['CI Name']=cis['CI Name'].astype(str).apply(lambda x: re.sub("\(|\)|\{|\}|\[|\]|\'|\"|\´|\»|\«|\/|\\\\", "",x))
					print('','CIs with Special Characteres (Auto fixed): '.upper() + str(np.shape(cis_chars)[0]),'-'*len('CIs with Special Characteres (Auto fixed):'),'',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				sites_reg=cis[['Site','Region','Site Group']].drop_duplicates()
				dup_sites_reg=sites_reg[sites_reg.iloc[:,0].duplicated(keep=False)].sort_values([sites_reg.columns[0]])
				###duplicated region per site in CIs
				cis_locations=[]
				if np.shape(dup_sites_reg)[0]>0:
					cis_locations.append(dup_sites_reg)
					print('','Duplicated Location per Site in CIs data: '.upper() +str(np.shape(dup_sites_reg)[0]),'-'*len('Duplicated Location per Site in CIs data:'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				##check DNS, CI Description and Domain
				filtered_cis=cis.filter(regex=re.compile('CI N',re.IGNORECASE))
				filtered_desc=cis.filter(regex=re.compile('DESC',re.IGNORECASE))
				filtered_region=cis.filter(regex=re.compile('REG',re.IGNORECASE))
				filtered_site_group=cis.filter(regex=re.compile('SITE GROUP',re.IGNORECASE))
				filtered_dns=cis.filter(regex=re.compile('DNS',re.IGNORECASE))      
				with_region=(cis[filtered_desc.columns[0]]==cis[filtered_cis.columns[0]] + ' | ' + cis[filtered_region.columns[0]]).all()
				with_sitegroup=(cis[filtered_desc.columns[0]]==cis[filtered_cis.columns[0]] + ' | ' + cis[filtered_site_group.columns[0]]).all()
				well_conc=[with_region,with_sitegroup]
				if well_conc[0]==False & well_conc[1]==False:
					cis['Suggested '+filtered_desc.columns[0]]=cis[filtered_cis.columns[0]] + ' | ' + cis[filtered_region.columns[0]]
					print('','CI Description concatenation:'.upper(),'-'*len('CI Description concatenation:'),'A concatenation with Region is suggested. Please check the Data Model','',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				
				well_conc_dns=(cis[filtered_dns.columns[0]]==cis[filtered_cis.columns[0]]).all()
				if well_conc_dns==False:
					cis['Suggested '+ filtered_dns.columns[0]]=cis[filtered_cis.columns[0]]
					print('','DNS Host Name values:'.upper(),'-'*len('DNS Host Name values:'),'Replaced by CI Name values','',sep='\n',file=open(report +'warnings'+names[j]+'.txt','a',encoding='utf8'))       
		
				else:
					None
		
				duplicate_cis=[]
				dup_cis=pd.DataFrame(cis[filtered_cis.duplicated(keep=False)].drop_duplicates()).sort_values([filtered_cis.columns[0]])
				duplicate_cis.append(dup_cis.groupby(filtered_cis.columns.values[0]).size().reset_index(name='counts'))
				if len(dup_cis)>0:
					print('','Duplicated CI Names (excluding duplicate rows): '.upper() + str(len(duplicate_cis)),'-'*len('Duplicated CI Names (excluding duplicate rows):'),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))
				else:
					None
				###check product catalogue
				opcat_template='Prod_Cats_V2'
				
				template=pd.read_csv(glob.glob(opcat_template+'/*.csv')[0],sep=',')
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
				wrong_catalogue.rename(columns={'FIELD':''},inplace=True)
				#wrong_catalogue=wrong_catalogue.set_index(wrong_catalogue.columns[0])
				if len(wrong_catalogue)>0:
					print('','WRONG VALUES IN PRODUCT CATEGORIZATION:','-'*len('WRONG VALUES IN PRODUCT CATEGORIZATION:'),wrong_catalogue.set_index(wrong_catalogue.columns[0]),'',sep='\n',file=open(report +'errors'+names[j]+'.txt','a',encoding='utf8'))              
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
				print('','CIs with non existing sites: '.upper()+str(np.shape(new_sites_in_cis)[0]),'-'*len('CIs with non existing sites:'),'',sep='\n',file=open(report +'errorsCIs.txt','a',encoding='utf8'))
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
				print('','Mismatched Locations between CIs and Sites data: '.upper()+ str(np.shape(wrong_locations)[0]),'-'*len('Mismatched Locations between CIs and Sites data:'),'',sep='\n',file=open(report +'errorsCIs.txt','a',encoding='utf8'))
			else:
				None
			
		else:
			None
	with pd.ExcelWriter(report + company + '_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
		if len(sites)>0:
			sites.to_excel(writer, 'sites',index=False)
			if np.shape(sites_chars)[0]>0:
				sites_chars.to_excel(writer,'Sites with special characteres',index=False)
			else:
				None
			if np.shape(new_sites)[0]>0:
				new_sites.to_excel(writer,'New Sites',index=False)
			else:
				None    		
			if np.shape(dup_sites)[0]>0:
				dup_sites.to_excel(writer, 'Duplicate Sites',index=False)
			else:
				None
			if len(existing_sites_list)>0:
				existing_sites_list[0].to_excel(writer, 'Existing Location',index=False)
			else:
				None
			if len(wrong_locations_sites_list)>0:
				wrong_locations_sites_list[0].to_excel(writer, 'Region Issues in Sites',index=False)
			else:
				None
		else:
			print('','No Sites data to validate'.upper(),sep='\n',file=open(report +'summary' + names[j] + 'Sites.txt','a',encoding='utf8'))
		if len(cis)>0:
			cis.to_excel(writer, 'cis',index=False)
			if np.shape(existing_cis)[0]>0:
				existing_cis.to_excel(writer,'CIs already in CMDB',index=False)
			else:
				None
			if np.shape(new_sites_in_cis)[0]>0:
				new_sites_in_cis.to_excel(writer,'CIs with non existing sites',index=False)
			else:
				None
			if len(cis_locations)>0:
				cis_locations[0].to_excel(writer, 'Region Issues in CIs',index=False)
			else:
				None
			if len(cis_sites_locations)>0:
				cis_sites_locations[0].to_excel(writer, 'CIs Sites Region Issues',index=False)
			else:
				None
			if np.shape(cis_chars)[0]>0:
				cis_chars.to_excel(writer,'CIs with special characteres',index=False)
			else:
				None   
			if np.shape(dup_cis)[0]>0:
				dup_cis.to_excel(writer, 'Duplicate CIs',index=False)
			else:
				None
			if len(wrong_Tiers)>0:
				wrong_Tiers.to_excel(writer,'wrong_tiers',index=False)
			else:
				None
			if len(prod_missing_final)>0:
				prod_missing_final.to_excel(writer,'wrong_product_name',index=False)

			else:
				None
			if len(cis_chars)>0:
				cis_chars.to_excel(writer,'CIs with special Characters',index=False)

			else:
				None
			  
		else:
			print('','No CIs data to validate'.upper(),sep='\n',file=open(report +'summary' + names[j] + 'CIs.txt','a',encoding='utf8'))
		writer.save()
	print(dt.datetime.now() - startTime)	