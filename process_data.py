import numpy as np
import pandas as pd
import re
import os
import glob
from difflib import get_close_matches
import itertools
import sys
import datetime as dt
from io import StringIO
import csv

def process_file(path,company,report,history):
	files=pd.DataFrame()
	sheets=[]
	cis=[]
	sites=[]
	all_sites=pd.DataFrame(columns=['Site','Region','Site Group','City'])
	for j in range(len(glob.glob(path+'/*'))):
		#for k in range(len(pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names)):
		#	sheets.append(pd.read_excel(glob.glob(path+'*')[j],pd.ExcelFile(glob.glob(path+'*')[j]).sheet_names[k]))     
    		files=pd.read_excel(glob.glob(path+'*')[j],sheet_name=None)
    		for frame in files.keys():
        		sheets.append(files[frame])

	sites_fields=['Company','Site Name','Site Alias','Description','Region','Site Group','Street','Country','City','Latitude','Longitude','Location ID','Additional Site Details','Maintenance Circle Name','Site Type','Status']
	cis_fields=['Company','CI Name','CI Description','Tag Number','System Role','Status','Priority','Additional Information','Tier 1','Tier 2','Tier 3','Product Name','Model Version','Manufacturer','Region','Site Group','Site','DNS Host Name','Domain','CI ID']
	fields=sites_fields+cis_fields
	char_num=[254,60,60,255,60,60,90,60,60,12,12,30,0,70,'-','-',254,254,254,64,30,'-','-',254,60,60,60,254,254,254,60,60,60,254,254,64]
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
			names.append('Sites Validation')
			unmatched_fields.append(list(set(sites_fields) - set(sheets[j])))
		if (sheets[j].columns.str.contains('CI N',case=False).any()):
			names.append('CIs Validation')
			unmatched_fields.append(list(set(cis_fields) - set(sheets[j])))   
		else:
			None

	if len(list(itertools.chain(*unmatched_fields)))>0:
		print('','#'*24,'#' +' DATA TO BE VALIDATED '+ '#','#'*24,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
		print('Missing or Mismatched fields:'.upper(),'',pd.DataFrame(pd.Series(list(itertools.chain(*unmatched_fields))).rename('Field')),'','Please use the templates for sites and CIs, provided in home page','',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
	else:
		all_fields=pd.concat([pd.Series(fields).rename('Field'),pd.Series(field_type).rename('Mandatory'),pd.Series(char_num).rename('Maximum allowed')],axis=1)
		common_fields=[]
		for j in range(len(sheets)):
			###first overview prints####
			print('','#'*20,str(j+1)+':' names[j].upper(),'#'*20,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))		
			if len(os.listdir(path))==1:
				print(os.listdir(path)[0],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			else:
				print(os.listdir(path)[j],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			blank_cases=[]
			common_fields.append(all_fields.merge(pd.DataFrame(pd.Series(sheets[j].columns).rename('Field')),on='Field',how='inner').drop_duplicates())
			count_chars=[]
			null_columns=sheets[j][sheets[j].columns[sheets[j].isnull().any()]]  		
			###blanks and char num
			for i in range(len(sheets[j].columns)):
				blank_find=sheets[j].iloc[:,i][sheets[j].iloc[:,i].astype(str).apply(lambda x: x[0].isspace() or x[len(x)-1].isspace())]
				if len(blank_find)>0:
					blank_cases.append(blank_find.iloc[0])
				else:
					blank_cases.append('None')
				#blank_cases.append(sheets[j].iloc[:,i][sheets[j].iloc[:,i].astype(str).apply(lambda x: x[0].isspace() or x[len(x)-1].isspace())].unique())
				count_chars.append(sheets[j].iloc[:,i].apply(lambda x: x if pd.isnull(x) else (len(str(round(x,5))) if type(x)==float else len(str(x)))).max())
				sheets[j].iloc[:,i]=sheets[j].iloc[:,i].apply(lambda x: x.strip() if type(x)==str else x)
			##count max number of characteres per field
			chars=pd.concat([pd.Series(sheets[j].columns).rename('Field'),pd.Series(count_chars).rename('Characters')],axis=1)
			b=common_fields[j].merge(chars,on='Field',how='outer')
			c=b.iloc[:,[0,1,3,2]]
			print('Number of records:'.upper(),'-'*len('Number of records:'), str(np.shape(sheets[j])[0]),'','Field Names and Maximum number of Characteres per field:'.upper(),'-'*len('Field Names and Maximum number of Characteres per field:'),c,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))			
			#if len(np.unique(sheets[j].filter(regex=re.compile('COMPANY',re.IGNORECASE))))>1:
			#	print('Company Name is wrongly filled: '.upper(),'-'*len('Company Name is wrongly filled: '),np.unique(sheets[j].filter(regex=re.compile('COMPANY',re.IGNORECASE))),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			#else:
			#	print('bla')
			#check for blank spaces
			#blanks=pd.concat([pd.Series(sheets[j].columns).rename('Field'),pd.Series(blank_cases).rename('Cases'),pd.Series(blank_cases).apply(lambda x: len(x)).rename('Count')],axis=1)
			blanks=pd.concat([pd.Series(sheets[j].columns).rename('Field'),pd.Series(blank_cases).apply(lambda x: len(x)).rename('Count'),pd.Series(blank_cases).rename('Case example'),],axis=1)
			if np.shape(blanks[blanks['Case example']!='None'])[0]>0:
				blank_spaces=blanks[blanks['Case example']!='None']
				print('Fields with blanks spaces: (Auto Removed)'.upper(),'-'*len('Fields with blanks spaces: (Auto Removed)'),blank_spaces,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))

			else:
				None
			
			if np.shape(null_columns.isnull().sum())[0]>0:
				null_fields=pd.DataFrame(null_columns.isnull().sum())
				null_fields.reset_index(level=0, inplace=True)
				null_fields.rename(columns={'index':'Field',0:'Count'},inplace=True)
				null_fields=null_fields.merge(all_fields.iloc[:,:-1],on='Field',how='inner').drop_duplicates()

				print('Fields with Null Values:'.upper(),'-'*len('Field with Null Values:'),null_fields,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
 
			else:
				None

			###duplicated rows
			if np.shape(sheets[j][sheets[j].duplicated()])[0]>0:
				dup_rows=np.shape(sheets[j][sheets[j].duplicated()])[0]
				print('Number of Duplicated rows: (Auto Removed)'.upper(),'-'*len('Number of Duplicated rows: (Auto Removed)'),dup_rows,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			else:
				None
			
			sheets[j].drop_duplicates(inplace=True)
			####locations
			sheets[j].fillna('',inplace=True)
			filtered_locations=sheets[j].filter(regex=re.compile('REG|GROUP|CITY',re.IGNORECASE))
			length=[]
			print('Location values:'.upper(),'-'*len('Location values:'),sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			for k in range(np.shape(filtered_locations)[1]):
				length.append(len(filtered_locations.iloc[:,k].unique()))
				a=pd.DataFrame(pd.Series(filtered_locations.iloc[:,k].unique()).sample(n=min(length), random_state=1))
				a.rename(columns={0:filtered_locations.columns[k]},inplace=True)
				print(a,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
				#print('-'*len('Location values:'),filtered_locations.columns[k],filtered_locations.iloc[:,k].unique(),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
				if filtered_locations.iloc[:,k].str.isupper().any():
					#sheets[j][filtered_locations.columns[k]]=filtered_locations.iloc[:,k].apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else x.title())
					print('Found Upper cases in ' + filtered_locations.columns[k],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
				else:
					None
			####ITSM sites
			#itsm_sites='/ITSM_sites'
			itsm_columns=['Site Name','Region','Site Group','City']
			print('All Existing Sites in CMDB:'.upper(),'-'*len('All Existing Sites in CMDB:'),sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			if len(glob.glob(history+'/*'))==0:
				sites_itsm=pd.DataFrame(columns=itsm_columns)
				sites_itsm=sites_itsm[itsm_columns]
				print('No CMDB Report: Suggest to get a full report of existing sites','',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			else:
				sites_itsm=pd.read_excel(glob.glob(history+'/*')[0],pd.ExcelFile(glob.glob(history+'/*')[0]).sheet_names[0])
				#sites_itsm=sites_itsm[sites_itsm['PrimAlias']==0]
				sites_itsm=sites_itsm.filter(regex=re.compile('SITE N|REG|GROUP|CITY',re.IGNORECASE))
				sites_itsm.rename(columns={
					sites_itsm.columns[0]:'Site Name',
					sites_itsm.columns[1]:'Region',
					sites_itsm.columns[2]:'Site Group',
					sites_itsm.columns[3]:'City'},inplace=True)
				#sites_itsm=sites_itsm[itsm_columns]
				print(np.shape(sites_itsm)[0],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))                  
			###########################################################################################################################
			#SITES            
			if (~sheets[j].columns.str.contains('CI N',case=False).any()):
				#existing_sites=pd.DataFrame()
				wrong_locations_sites_list=[]
				if np.shape(sites_itsm)[0]>0:
					all_sites=sheets[j].merge(sites_itsm.set_index(sites_itsm.columns[0]),left_on=sheets[j].filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='outer',indicator=True).drop_duplicates()
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
						print('Sites to upload already existing in CMDB: '.upper(),'-'*len('Sites to upload already existing in CMDB: '),'#: ' + str(np.shape(existing_sites)[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
						print('New sites to upload in CMDB: '.upper(),'-'*len('New Sites to upload in CMDB: '),'#: ' + str(np.shape(new_sites)[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))

						wrong_locations_sites=pd.concat([existing_sites[(existing_sites['Region in Sites'])!=(existing_sites['Region in ITSM Sites'])],existing_sites[(existing_sites['Site Group in Sites'])!=(existing_sites['Site Group in ITSM Sites'])]],axis=0).drop_duplicates()
						if np.shape(wrong_locations_sites)[0]>0:
							print('Existing sites with mismatched locations:'.upper(),'-'*len('Existing sites with mismateched locations:'),'#: ' +str(np.shape(wrong_locations_sites)[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
							wrong_locations_sites_list.append(wrong_locations_sites)
						else:
							None
					else:
						None
				else:
					None
				##new_sites###
				sites.append(sheets[j])
				sites[0]['Longitude']=sites[0].filter(regex=re.compile('Lon',re.IGNORECASE)).iloc[:,0].astype(str).apply(lambda x: re.findall(r'-?\d+\.?\d*', re.sub('[^A-Za-z0-9]+', '.',x))).apply(lambda x: ''.join(x) if len(x)>0 else 'NaN').astype(float)
				sites[0]['Latitude']=sites[0].filter(regex=re.compile('Lat',re.IGNORECASE)).iloc[:,0].astype(str).apply(lambda x: re.findall(r'-?\d+\.?\d*', re.sub('[^A-Za-z0-9]+', '.',x))).apply(lambda x: ''.join(x) if len(x)>0 else 'NaN').astype(float)
				####duplicated sites
				sites_locations=pd.concat([sheets[j].filter(regex=re.compile('SITE N',re.IGNORECASE)),sheets[j].filter(regex=re.compile('REG|GROUP|CITY',re.IGNORECASE))],axis=1)
				dup_sites=pd.DataFrame(sheets[j][pd.DataFrame(sites_locations.iloc[:,0]).duplicated(keep=False)].drop_duplicates()).sort_values([sites_locations.columns[0]])
				duplicate_sites=[]
				duplicate_sites.append(dup_sites.groupby(pd.DataFrame(sites_locations.iloc[:,0]).columns.values[0]).size().reset_index(name='counts'))
				if np.shape(duplicate_sites[0])[0]>0:
					print('Duplicated Site Names (excluding duplicate rows):'.upper(),'-'*len('Duplicated Site Names (excluding duplicate rows):'),duplicate_sites[0],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))  
				else:
					None
				#sites[0].to_csv(report + company +'sites_to_add'+'.csv')
			else:
				None
			#CIS
			if sheets[j].columns.str.contains('CI N',case=False).any():
				cis_locations=[]
				cis.append(sheets[j])
				cis_chars=pd.DataFrame(cis[0][cis[0]['CI Name'].astype(str).str.contains("\"|\'|´",regex=True)].drop_duplicates())
				if np.shape(cis_chars)[0]>0:
					cis[0]['CI Name']=cis[0]['CI Name'].astype(str).apply(lambda x: re.sub("\'|\"|\´", "",x))
					print('CIs with quotes (Quotes Auto Removed):'.upper(),'-'*len('CIs with quotes (Auto Removed'),'#: ' + str(np.shape(cis_chars)[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
				else:
					None
				#site_name=cis[0][cis[0].columns[~cis[0].columns.str.contains('Group',case=False)].tolist()].filter(regex=re.compile('SITE',re.IGNORECASE))
				#sites_reg=pd.concat([site_name,cis[0].filter(regex=re.compile('REG|SITE GROUP',re.IGNORECASE))],axis=1).drop_duplicates()
				sites_reg=cis[0][['Site','Region','Site Group']].drop_duplicates()
				#sites_reg=cis[0].filter(regex=re.compile('SITE N|SITE+|SITE*|REG|GROUP',re.IGNORECASE)).drop_duplicates()#merge(all_sites2.filter(regex=re.compile('SITE N|SITE+|SITE*|REG|GROUP|CITY',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=cis[0].filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='inner',indicator=True).drop_duplicates()
				dup_sites_reg=sites_reg[sites_reg.iloc[:,0].duplicated(keep=False)].sort_values([sites_reg.columns[0]])
				if np.shape(dup_sites_reg)[0]>0:
					cis_locations.append(dup_sites_reg)
					print('Duplicated Location per Site in CIs data:'.upper(),'-'*len('Duplicated Location per Site in CIs data:'),'#: ' +str(np.shape(dup_sites_reg)[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
					#print(np.shape(dup_sites_reg)[0],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
				else:
					None
					#print('Site with correct region related:','',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
				filtered_cis=cis[0].filter(regex=re.compile('CI N',re.IGNORECASE))
				filtered_desc=cis[0].filter(regex=re.compile('DESC',re.IGNORECASE))
				filtered_region=cis[0].filter(regex=re.compile('REG',re.IGNORECASE))
				filtered_site_group=cis[0].filter(regex=re.compile('SITE GROUP',re.IGNORECASE))
				filtered_dns=cis[0].filter(regex=re.compile('DNS',re.IGNORECASE))      
				#if (np.shape(filtered_desc)[1]>0) & ((np.shape(filtered_region)[1]>0) | (np.shape(filtered_site_group)[1]>0)):
				with_region=(cis[0][filtered_desc.columns[0]]==cis[0][filtered_cis.columns[0]] + ' | ' + cis[0][filtered_region.columns[0]]).all()
				with_sitegroup=(cis[0][filtered_desc.columns[0]]==cis[0][filtered_cis.columns[0]] + ' | ' + cis[0][filtered_site_group.columns[0]]).all()
				well_conc=[with_region,with_sitegroup]
				if well_conc[0]==False & well_conc[1]==False:
					cis[0]['Suggested '+filtered_desc.columns[0]]=cis[0][filtered_cis.columns[0]] + ' | ' + cis[0][filtered_region.columns[0]]
					print('CI Description concatenation:'.upper(),'-'*len('CI Description concatenation:'),'A concatenation with Region is suggested. Please check the Data Model','',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))

				else:
					None
				
				well_conc_dns=(cis[0][filtered_dns.columns[0]]==cis[0][filtered_cis.columns[0]]).all()
				if well_conc_dns==False:
					cis[0]['Suggested '+ filtered_dns.columns[0]]=cis[0][filtered_cis.columns[0]]
					print('DNS Host Name values:'.upper(),'-'*len('DNS Host Name values:'),'Replaced by CI Name values','',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))       
		
				else:
					None
		
				duplicate_cis=[]
				dup_cis=pd.DataFrame(cis[0][filtered_cis.duplicated(keep=False)].drop_duplicates()).sort_values([filtered_cis.columns[0]])
				duplicate_cis.append(dup_cis.groupby(filtered_cis.columns.values[0]).size().reset_index(name='counts'))
				if np.shape(duplicate_cis[0])[0]>0:
					print('Duplicated CI Names (excluding duplicate rows):'.upper(),'-'*len('Duplicated CI Names (excluding duplicate rows):'),'#: ' + str(np.shape(duplicate_cis[0])[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))

				else:
					None

				#cis_chars=pd.DataFrame(cis[0][filtered_cis.iloc[:,0].astype(str).str.contains("\;|\.|\,|\:|\$|\s",regex=True)].drop_duplicates())
				
				#if cis[0].filter(regex=re.compile('PRIORITY',re.IGNORECASE)).isnull().any():
				#	print('Priority not correctly filled')
				#else:
				#	prior=pd.DataFrame(np.unique(cis[0].filter(regex=re.compile('PRIORITY',re.IGNORECASE))))
				#	prior.rename(columns={0:'Values'},inplace=True)
				#	print('PRIORITIES USED: ','-'*len('PRIORITIES USED: '),prior,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))

				#########################
				##product categorization
				#data=urlopen("https://fiespmidup1.int.net.nokia.com:8443/customer_onboarding/csv/oneitsm_ProdCats.csv").read().decode('ascii','ignore')
				#datafile=StringIO(data)
				#csvReader=csv.reader(datafile)
				#prod_cat_list=[]
				#for row in csvReader:
				#	prod_cat_list.append(row)
				#prod_cat=pd.DataFrame(prod_cat_list)
				#headers = prod_cat.iloc[0]
				#prod_cat  = pd.DataFrame(prod_cat.values[1:], columns=headers)
				#template=prod_cat
				
				opcat_template='Prod_Cats'
				#template=pd.read_excel(glob.glob(opcat_template+'/*')[0],pd.ExcelFile(glob.glob(opcat_template+'/*')[0]).sheet_names[5])
				template=pd.read_csv(glob.glob(opcat_template+'/*')[0],sep=',')
				template.rename(columns={template.columns[0]:'Correct Tier 1',
							 template.columns[1]:'Correct Tier 2',
							 template.columns[2]:'Correct Tier 3',
							 template.columns[4]:'Correct Manufacturer'},
						inplace=True)
				#template.rename(columns={template.columns[3]:'Product Name'},inplace=True)
				prodcats_cis=cis[0][['Tier 1','Tier 2','Tier 3','Product Name','Manufacturer']]
				#prodcats_cis=cis[0].filter(regex=re.compile('TIER|PRODUCT N|MANUF',re.IGNORECASE))
				prod_missing=prodcats_cis.loc[~prodcats_cis['Product Name'].isin(template.iloc[:,3])].drop_duplicates()
				prod_name=prod_missing.filter(regex=re.compile('Product N',re.IGNORECASE)).iloc[:,0]
				prod_match=prod_name.apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else get_close_matches(x,template.iloc[:,3].astype(str).unique().tolist()))
				prod_suggested=prod_name.apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else get_close_matches(x,template.iloc[:,3].astype(str).unique().tolist(),1)).apply(lambda x: x if pd.isnull(x) or type(x)==float or type(x)==int else ''.join(x))
				prod_missing_final=pd.concat([prod_missing,prod_suggested.rename('Suggested Product Name'),prod_match.rename('Others PN match')],axis=1)
				#wrong  product names tier classification
				same_prod=prodcats_cis[~prodcats_cis['Product Name'].isna()].merge(template,left_on=prod_name.name,right_on='Product Name',how='inner').drop_duplicates()
				wrongcats1=same_prod.iloc[:,[0,5,3,4,8]][(same_prod.iloc[:,0]!=same_prod.iloc[:,5]) & (same_prod.iloc[:,1]==same_prod.iloc[:,6]) & (same_prod.iloc[:,2]==same_prod.iloc[:,7])].drop_duplicates()
				wrongcats2=same_prod.iloc[:,[1,6,3,4,8]][(same_prod.iloc[:,1]!=same_prod.iloc[:,6]) & (same_prod.iloc[:,0]==same_prod.iloc[:,5]) & (same_prod.iloc[:,2]==same_prod.iloc[:,7])].drop_duplicates()
				wrongcats3=same_prod.iloc[:,[2,7,3,4,8]][(same_prod.iloc[:,2]!=same_prod.iloc[:,7]) &(same_prod.iloc[:,0]==same_prod.iloc[:,5]) & (same_prod.iloc[:,1]==same_prod.iloc[:,6])].drop_duplicates()
				wrongcats1and2=same_prod.iloc[:,[0,5,1,6,3,4,8]][(same_prod.iloc[:,0]!=same_prod.iloc[:,5]) & (same_prod.iloc[:,1]!=same_prod.iloc[:,6]) & (same_prod.iloc[:,2]==same_prod.iloc[:,7])].drop_duplicates()
				wrongcats1and3=same_prod.iloc[:,[0,5,2,7,3,4,8]][(same_prod.iloc[:,0]!=same_prod.iloc[:,5]) & (same_prod.iloc[:,1]==same_prod.iloc[:,6]) & (same_prod.iloc[:,2]!=same_prod.iloc[:,7])].drop_duplicates()
				wrongcats2and3=same_prod.iloc[:,[1,6,2,7,3,4,8]][(same_prod.iloc[:,1]!=same_prod.iloc[:,6]) & (same_prod.iloc[:,2]!=same_prod.iloc[:,7]) & (same_prod.iloc[:,0]==same_prod.iloc[:,5])].drop_duplicates()
				wrongcats1and2and3=same_prod.iloc[:,[0,5,1,6,2,7,3,4,8]][(same_prod.iloc[:,0]!=same_prod.iloc[:,5]) & (same_prod.iloc[:,1]!=same_prod.iloc[:,6]) & (same_prod.iloc[:,2]!=same_prod.iloc[:,7])].drop_duplicates()
				wrong_manufacturer=same_prod.iloc[:,[3,4,8]][(same_prod.iloc[:,4]!=same_prod.iloc[:,8])].drop_duplicates()                        
				cis_list=[]
				cis_list.append(prod_missing_final)
				cis_list.append(wrongcats1)
				cis_list.append(wrongcats2)
				cis_list.append(wrongcats3)
				cis_list.append(wrongcats1and2)
				cis_list.append(wrongcats1and3)
				cis_list.append(wrongcats2and3)
				cis_list.append(wrongcats1and2and3)
				cis_list.append(wrong_manufacturer)
				issues_names=['Wrong product Name','Wrong Tier 1','Wrong Tier 2','Wrong Tier 3',
							  'Wrong Tier 1 and 2','Wrong Tier 1 and 3','Wrong Tier 2 and 3','Wrong Tier 1, 2 and 3',
							  'Wrong Manufacturer']
				issues=pd.concat([pd.Series(issues_names).rename('Field'),pd.Series(cis_list).apply(lambda x: len(x)).rename('COUNT')],axis=1)
				print('PRODUCT CATALOG ISSUES:','-'*len('PRODUCT CATALOG ISSUES:'),issues,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))               
				#cis[0].to_csv(report + company +'cis_to_add_'+'.csv')			
			else:
				None				
		cis_sites_locations=[]
		if (len(cis)>0):
			if (len(sites)>0):
				conc_sites=[]
				for i in range(len(sites_itsm.columns)):
					conc_sites.append(pd.concat([sites_itsm.iloc[:,i],sites_locations.iloc[:,i]],axis=0))
				all_sites2=pd.concat(conc_sites,axis=1)
				all_sites2.rename(columns={
				all_sites2.columns[0]:'Site Name',
				all_sites2.columns[1]:'Region',
				all_sites2.columns[2]:'Site Group',
				all_sites2.columns[3]:'City'},inplace=True)
			else:
				all_sites2=sites_itsm
			##CIs with non existing sites
			new_sites_in_cis=sites_reg.merge(all_sites2.filter(regex=re.compile('SITE N|SITE+|SITE*|REG|GROUP|CITY',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=sites_reg.filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='outer',indicator=True).drop_duplicates()
			new_sites_in_cis=new_sites_in_cis[new_sites_in_cis['_merge']=='left_only'].iloc[:,:-1]
			if np.shape(new_sites_in_cis)[0]>0:
				new_sites_in_cis.rename(columns={new_sites_in_cis.columns[1]:'Region',
								 new_sites_in_cis.columns[2]:'Site Group'},
							inplace=True
						       )
				new_sites_in_cis=new_sites_in_cis[['Site','Region','Site Group']]
				print('CIs with non existing sites'.upper(),'-'*len('CIs with non existing sites'),np.shape(new_sites_in_cis)[0],'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			else:
				None
			existing_sites2=sites_reg.merge(all_sites2.filter(regex=re.compile('SITE N|REG|SITE GROUP|CITY',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=sites_reg.filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='inner').drop_duplicates()
			
			#existing_sites2=cis[0].filter(regex=re.compile('SITE N|SITE+|SITE*|REG|GROUP',re.IGNORECASE)).merge(all_sites2.filter(regex=re.compile('SITE N|SITE+|SITE*|REG|GROUP|CITY',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=cis[0].filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='inner').drop_duplicates()
			#new_sites=sites_reg.merge(all_sites2.filter(regex=re.compile('SITE N|REG|GROUP|CITY',re.IGNORECASE)).set_index(all_sites2.columns[0]),left_on=sites_reg.filter(regex=re.compile('SITE',re.IGNORECASE)).columns[0],right_index=True,how='outer',indicator=True).drop_duplicates()
			#new_sites=new_sites[new_sites['_merge']=='left_only']
			#if np.shape(all_sites2)[0]>0:
			#print('Sites in CIS not found in CMDB: '.upper(),'-'*len('Sites in CIS not found in CMDB: '),new_sites,'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			
			existing_sites2.rename(columns={
				existing_sites2.filter(regex=re.compile('REGION',re.IGNORECASE)).columns[0]:'Region in CIs',
				existing_sites2.filter(regex=re.compile('REGION',re.IGNORECASE)).columns[1]:'Region in Sites',
				existing_sites2.filter(regex=re.compile('SITE GROUP',re.IGNORECASE)).columns[0]:'Site Group in CIs',
				existing_sites2.filter(regex=re.compile('SITE GROUP',re.IGNORECASE)).columns[1]:'Site Group in Sites'},inplace=True)

			wrong_locations=pd.concat([existing_sites2[(existing_sites2['Region in CIs'])!=(existing_sites2['Region in Sites'])],existing_sites2[(existing_sites2['Site Group in CIs'])!=(existing_sites2['Site Group in Sites'])]],axis=0).drop_duplicates()

			if np.shape(wrong_locations)[0]>0:
				cis_sites_locations.append(wrong_locations)
				print('Mismatched Locations between CIs and Sites data:'.upper(),'-'*len('Mismatched Locations between CIs and Sites data:'),'#: '+ str(np.shape(wrong_locations)[0]),'',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))

			else:
				None
			
		else:
			None
	
	print('','#################','#Report Overview#'.upper(),'#################','',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
	
	with pd.ExcelWriter(report + company + '_issues_report_'+ dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S") +'.xlsx',engine='xlsxwriter') as writer:
		if len(sites)>0:
			sites[0].to_excel(writer, 'sites',index=False)			
			if np.shape(dup_sites)[0]>0:
				dup_sites.to_excel(writer, 'Duplicate Sites',index=False)
			else:
				print('No Duplicate Sites',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			#if np.shape(existing_sites)[0]>0:
			#	existing_sites.to_excel(writer, 'Existing Sites',index=False)
			#else:
			#	print(' No existing Sites',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			if len(wrong_locations_sites_list)>0:
				wrong_locations_sites_list[0].to_excel(writer, 'Region Issues in Sites',index=False)
			else:
				print('No Wrong Locations in sites',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
		else:
			print('No Sites',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
		if len(cis)>0:
		#	try:
			cis[0].to_excel(writer, 'cis',index=False)
		#	except Exception as e:
		#		cis[0].to_csv(report + company +'cis_report_'+'.csv')	
			if np.shape(new_sites_in_cis)[0]>0:
				new_sites_in_cis.to_excel(writer,'CIs with non existing sites',index=False)
			if len(cis_locations)>0:
				cis_locations[0].to_excel(writer, 'Region Issues in CIs',index=False)
			else:
				print('No Wrong Locations in CIs',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			if len(cis_sites_locations)>0:
				cis_sites_locations[0].to_excel(writer, 'CIs Sites Region Issues',index=False)   
			else:
				print('No Wrong locations between sites and cis',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))    
			if np.shape(dup_cis)[0]>0:
				dup_cis.to_excel(writer, 'Duplicate CIs',index=False)
			else:
				print('No Duplicate CIs',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
			for i in range(len(cis_list)):
				if np.shape(cis_list[i])[0]>0:
					cis_list[i].to_excel(writer, issues_names[i],index=False)
				else:
					print('No ' + issues_names[i],sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))                       
		else:
			print('No CIs',sep='\n',file=open(report +'issues.txt','a',encoding='utf8'))
		writer.save()
