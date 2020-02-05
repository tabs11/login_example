#!/usr/bin/env python
from flask import Flask, request, redirect, url_for, render_template, send_from_directory,send_file,flash,session
from werkzeug.utils import secure_filename
from flask_login import LoginManager
from flask_login import UserMixin # subclass of flask user
from flask_login import login_required
from flask_login import login_user
from flask_login import logout_user
import shutil
import pandas as pd
import uuid
import os
import datetime as dt 
from datetime import datetime
import process_data
import process_noam_data
import process_res_cats
import process_site_history
import process_zte
import update_priority as up_prio
class User(UserMixin):
	def __init__(self, username,password):
		super(User, self).__init__()
		self.username = username
		self.password = password
	def is_active(self):
		return True
	def is_anonymous(self):
		return False
	def get_id(self):
		return self.username

# USER DATABASE
USERS = { # dictionary (username, User)
	'numartin' : User('numartin','pass'),
	'jschuur' : User('jschuur','teste'),
	'nawaz' : User('nawaz','nawaz'),
	'mccavitt' : User('mccavitt','mccavitt'),
	'abociat' : User('abociat','abociat'),
	'mariusgc' : User('mariusgc','mariusgc'),
	'gelei' : User('gelei','gelei'),
	'abociat' : User('abociat','abociat'),
	'marcinkl' : User('marcinkl','marcinkl'),
	'mayayada' : User('mayayada','mayayada'),
	'dees' : User('dees','dees'),
	'singaram' : User('singaram','singaram'),
	'nvp' : User('nvp','nvp'),
	'avilcean' : User('avilcean','avilcean'),
	'paulof'   : User('paulof','paulof'),
	'adriani'  : User('adriani','adriani'),
	'paagrawa'  : User('paagrawa','paagrawa'),
	'paagrawa_admin'  : User('paagrawa_admin','paagrawa_admin'),
	'mlacan'  : User('mlacan','mlacan'),
	'rserban' : User('rserban','rserban'),
	'matheka' : User('matheka','matheka'),
	'fjoseph' : User('fjoseph','fjoseph'),
	'difrance' : User('difrance','difrance')
	


	
}

application = Flask(__name__)


SECRET_KEY='bla'#str(uuid.uuid1())
application.secret_key = SECRET_KEY
CMDB_FOLDER = 'CMDB_templates/'
application.config['CMDB_FOLDER']=CMDB_FOLDER
application.config['ALLOWED_EXTENSIONS'] = set(['xlsx','xls','csv'])


@application.route('/test', methods=['GET'])
def index():
	
	return redirect("/", code=302)


# login views
@application.route('/login', methods=['GET'])
def login_get():
	return render_template('login.html')

@application.route('/login', methods=['POST'])
def login_post():
	# get details from post request
	username = request.form['username']
	password = request.form['password']
	# get  user
	try:
		user = USERS[username]
		session['username']=str(username)
	except KeyError:
		user = None
	# validate user
	if user and user.password == password:
		login_user(user)
		if request.args.get("next"):
			return redirect(request.args.get("next"))
		else:
			return redirect('/')
	else:
		flash("wrong username or password")
		return render_template('login.html')



def allowed_file(filename):
	return '.' in filename and \
		   filename.rsplit('.', 1)[1] in application.config['ALLOWED_EXTENSIONS']


##########################################################################
#####download templates
@application.route('/return-file/')
@login_required
def return_file():
	filename='Templates.zip'
	return send_file(os.path.join(application.config['CMDB_FOLDER'])+filename,attachment_filename=filename, as_attachment=True)


@application.route('/test')
@login_required
def file_downloads():
	return render_template('home.html')


@application.route('/')
def drop():
	return render_template(
		'home.html',
		data=['Dummy Company','ALTAN Mexico','AT&T US','Airtel Chad','Airtel Congo B','Airtel Gabon','Airtel KE','Airtel Kenya','Airtel Madagascar','Airtel Malawi','Airtel Niger','Airtel Seychelles','Airtel Tanzania','Airtel Uganda','Airtel Zambia','Avantel CO','BHI NI','Bharti India','Capita TfL GB','Chorus NZ','Deutsche Telekom EAN DE','EMTS NG','Feenix NZ','IIJ Japan','ISAT EAN DE','MTN SIG HUB','NESC','Nexera PL','Nokia AVA','Optus AU','Orange Burkina Faso','Orange CALICO FR','Orange RIP FR','Rakuten JP','S-Bahn Berlin DE','T-Mobile US','TTN DK','Telenor DK','Telenor PK','Telia DK','Three Ireland','Vodacom TZ','Vodacom ZA','Vodacom ZA DWDM','Vodafone QA','Wing EU','Wing EU ATT US','Wing EU Marubeni JP','Wing EU TELE2','ZEOP RE'])



@application.route("/test" , methods=['GET', 'POST'])
def home():
	select = request.form.get('comp_select')
	if request.method == 'POST':
		user=session['username']
		session['company']=str(select)
		session['filename']=session['company']+'_'+str(uuid.uuid1())
		ID_FOLDER=session['filename']
		UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
		os.makedirs(ID_FOLDER)
		os.makedirs(UPLOAD_FOLDER)
			#msg = 'Successfull'
		if user!='numartin' and user!='paulof':
			print(user)
			return render_template('multi_upload_index.html')
		else:
			return render_template('cmdb_validation_admin.html')



##############################################
####CMDB inventory
@application.route('/site_data', methods=['GET','POST'])
@login_required
def site_data():
	session['filename']=session['company']+'_'+str(uuid.uuid1())
	SITE_FOLDER=session['filename']
	SITE_REPORT=SITE_FOLDER +'/Report/'
	os.makedirs(SITE_FOLDER)
	os.makedirs(SITE_REPORT)
	return render_template('site_index.html')


@application.route('/site_upload', methods=['GET'])
@login_required
def site_upload():
	
	SITE_FOLDER=session['filename']
	msg=SITE_FOLDER.split('_')[0]
	SITE_REPORT=SITE_FOLDER +'/Report/'
	process_site_history.sites_cis_report(company=SITE_FOLDER.split('_')[0],site_report=SITE_REPORT)
	site_filenames=os.listdir(SITE_REPORT)
	return render_template('site_upload.html', site_filenames=site_filenames,mgs=msg)

@application.route('/site_report/<filename>')
@login_required
def uploaded_site_file(filename):
	SITE_FOLDER=session['filename']
	SITE_REPORT=SITE_FOLDER +'/Report/'
	return send_from_directory(SITE_REPORT,filename)

######################################################################

#####input files to validate


@application.route('/data', methods=['GET','POST'])
def data_to_validate():
	ID_FOLDER=session['filename']
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
	msg3=None
	uploaded_files = request.files.getlist("file[]")
	for file in uploaded_files:
		# Check if the file is one of the allowed types/extensions
		if file and allowed_file(file.filename):
			
			# Make the filename safe, remove unsupported chars
			filename = secure_filename(file.filename)

			# Move the file form the temporal folder to the upload
			file.save(os.path.join(UPLOAD_FOLDER, filename))
			filenames=os.listdir(UPLOAD_FOLDER)
			msg3=filenames
		else:
			msg3='Please select a valid extension (.xls or .xlsx)'

	return render_template('multi_upload_index.html',msg3=msg3)

@application.route('/upload', methods=['POST'])
#@cache.cached(timeout=500)
@login_required
def upload():
	msg_company=None
	msg_cmdb1=None
	msg_cmdb2=None
	msgCIs=None
	msgSites=None
	msg=None
	msg2=None
	msg3=None
	msg4=None
	msg5=None
	msg6=None
	msg7=None
	msg8=None
	msg9=None
	msg10=None
	msg11=None
	msg12=None
	msg13=None
	#msg14=None
	msg15=None
	#msg16=None
	msg17=None
	msg18=None
	ID_FOLDER=session['filename']
	#ITSM_FOLDER=ID_FOLDER + '/ITSM_sites/'
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	os.makedirs(DOWNLOAD_FOLDER)
	if len(os.listdir(UPLOAD_FOLDER))>0:
		#process_data.process_file(path=UPLOAD_FOLDER,company=ID_FOLDER.split('_')[0],report=DOWNLOAD_FOLDER,history=ITSM_FOLDER)
		#process_data_v3.open_input_file(path=UPLOAD_FOLDER,report=DOWNLOAD_FOLDER)
		process_data.process_file(path=UPLOAD_FOLDER,company=ID_FOLDER.split('_')[0],report=DOWNLOAD_FOLDER)
		
		#process_data_v2.check_correct_fields(report=DOWNLOAD_FOLDER)
		#process_data_v2.common_validation(path=UPLOAD_FOLDER,report=DOWNLOAD_FOLDER)
		#process_data_v2.specific_validation(report=DOWNLOAD_FOLDER,history=ITSM_FOLDER)
		filenames = [f for f in os.listdir(DOWNLOAD_FOLDER) if f.endswith(('.xlsx','csv'))]
		content_cmdb=''
		content_errors_Cis=''
		content_warnings_Cis=''
		content_summary_Cis=''	
		content_errors_Sites=''
		content_warnings_Sites=''
		content_summary_Sites=''
		msg_company=ID_FOLDER.split('_')[0]
		if 'Mismatched_fields.txt' in os.listdir(DOWNLOAD_FOLDER):
			text_mis_fields=open(DOWNLOAD_FOLDER+'Mismatched_fields.txt', 'r+',encoding='utf8')
			content_mis_fields = text_mis_fields.read()
			text_mis_fields.close()
			msg17='WRONG TEMPLATE USED OR FIELDS ARE MISSING FROM ORIGINAL TEMPLATE.'
			msg18='PLEASE USE THE CORRECT TEMPLATES PROVIDED IN HOMEPAGE.'
		else:
			content_mis_fields=''
			if 'summary_CMDB.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_cmdb=open(DOWNLOAD_FOLDER+'summary_CMDB.txt', 'r+',encoding='utf8')
				content_cmdb = text_cmdb.read()
				text_cmdb.close()
				msg_cmdb1='CMDB size: '.upper()
			else:
				msg_cmdb2='Missing CMDB inventory'.upper()
			msgCIs="CI'S VALIDATION:"
			if 'errorsCIs.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_errors_Cis=open(DOWNLOAD_FOLDER+'errorsCIs.txt', 'r+',encoding='utf8')
				content_errors_Cis = text_errors_Cis.read()
				text_errors_Cis.close()
				if len(content_errors_Cis)>1:
					msg='Errors found in CIs data. Please download the report and check the sheets.'
				else:
					msg2='Your CIs data has no errors and is ready to be uploaded.'
			else:
				msg3='CIS FILE NOT FOUND.'
	
			if 'warningsCIs.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_warnings_Cis=open(DOWNLOAD_FOLDER+'warningsCIs.txt', 'r+',encoding='utf8')
				content_warnings_Cis = text_warnings_Cis.read()
				text_warnings_Cis.close()
				if len(content_warnings_Cis)>1:
					msg7='WARNINGS:'
				else:
					content_warnings_Cis=''
					msg8='YOUR CIS DATA HAS NO WARNINGS.'
			else:
				None
	
			if 'summaryCIs.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_summary_Cis=open(DOWNLOAD_FOLDER+'summaryCIs.txt', 'r+',encoding='utf8')
				content_summary_Cis = text_summary_Cis.read()
				text_summary_Cis.close()
				msg15='SUMMARY:'
			else:
				None

			msgSites="SITES VALIDATION"
			if 'errorsSites.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_errors_Sites=open(DOWNLOAD_FOLDER+'errorsSites.txt', 'r+',encoding='utf8')
				content_errors_Sites = text_errors_Sites.read()
				text_errors_Sites.close()
				if len(content_errors_Sites)>1:
					msg4='Errors found in Sites data. Please download the report and check the sheets.'
				else:
					msg5='Your Sites data has no errors and is ready to be uploaded.'
			else:
				msg6='SITES FILE NOT FOUND'
	
			
			if 'warningsSites.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_warnings_Sites=open(DOWNLOAD_FOLDER+'warningsSites.txt', 'r+',encoding='utf8')
				content_warnings_Sites = text_warnings_Sites.read()
				text_warnings_Sites.close()
				if len(content_warnings_Sites)>1:
					msg10='WARNINGS:'
				else:
					msg11='YOUR SITES DATA HAS NO WARNINGS.'
			else:
				None
				
	
			if 'summarySites.txt' in os.listdir(DOWNLOAD_FOLDER):
				text_summary_Sites=open(DOWNLOAD_FOLDER+'summarySites.txt', 'r+',encoding='utf8')
				content_summary_Sites = text_summary_Sites.read()
				text_summary_Sites.close()
				msg13='SUMMARY:'
			else:
				None
	return render_template('multi_files_upload.html', 
		filenames=filenames,
		text_cmdb=content_cmdb,
		text_mis_fields=content_mis_fields,
		text_errors_Cis=content_errors_Cis,
		text_errors_Sites=content_errors_Sites,
		text_warnings_Cis=content_warnings_Cis,
		text_warnings_Sites=content_warnings_Sites,
		text_summary_Cis=content_summary_Cis,
		text_summary_Sites=content_summary_Sites,
		msg_company=msg_company,msg_cmdb1=msg_cmdb1,msg_cmdb2=msg_cmdb2,msgCIs=msgCIs,msgSites=msgSites,msg=msg,msg2=msg2,msg3=msg3,msg4=msg4,msg5=msg5,msg6=msg6,msg7=msg7,
		msg8=msg8,msg9=msg9,msg10=msg10,msg11=msg11,msg12=msg12,msg13=msg13,msg15=msg15,msg17=msg17,msg18=msg18)


@application.route('/report/<filename>')
@login_required
def uploaded_file(filename):
	ID_FOLDER=session['filename']
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	return send_from_directory(DOWNLOAD_FOLDER,filename)


##################################################################
################convert to NOAM

@application.route('/noam_data', methods=['GET','POST'])
def noam_data():
	msg3=None
	session['filename']=session['company']+'_'+str(uuid.uuid1())
	NOAM_FOLDER=session['filename']
	NOAM_UPLOAD=NOAM_FOLDER+'/NOAM_files/'
	NOAM_REPORT=NOAM_FOLDER +'/Report/'
	
	# Get the name of the uploaded files
	uploaded_files = request.files.getlist("file[]")
	for file in uploaded_files:
		# Check if the file is one of the allowed types/extensions
		if file and allowed_file(file.filename):
			# Make the filename safe, remove unsupported chars
			filename = secure_filename(file.filename)
			if not os.path.exists(NOAM_FOLDER):
				os.makedirs(NOAM_FOLDER)
				os.makedirs(NOAM_UPLOAD)
				os.makedirs(NOAM_REPORT)
			# Move the file form the temporal folder to the upload
			
			file.save(os.path.join(NOAM_UPLOAD, filename))
			filenames=os.listdir(NOAM_UPLOAD)
			msg3=filenames
		else:
			msg3='Please select a valid extension (.xls or .xlsx)'
	return render_template('index_NOAM_company.html',msg3=msg3)


@application.route('/noam_upload', methods=['GET'])
def noam_upload():
	NOAM_FOLDER=session['filename']
	NOAM_UPLOAD=NOAM_FOLDER+'/NOAM_files/'
	NOAM_REPORT=NOAM_FOLDER +'/Report/'
	# Get the name of the uploaded files
	if len(os.listdir(NOAM_UPLOAD))>0:
		process_noam_data.noam_files(file_path=NOAM_UPLOAD,company=NOAM_FOLDER.split('_')[0],NOAM_report=NOAM_REPORT)
		noam_filenames=os.listdir(NOAM_REPORT)

	return render_template('noam_files_upload.html', noam_filenames=noam_filenames)


@application.route('/NOAM_Report/<filename>')
def uploaded_NOAM_file(filename):
	NOAM_FOLDER=session['filename']
	NOAM_REPORT=NOAM_FOLDER +'/Report/'
	return send_from_directory(NOAM_REPORT,filename)

##########################################################
###RESOLUTION AND OPERATIONAL Categories valdidation
@application.route('/op_res_cats', methods=['GET','POST'])
@login_required
def op_res_cats_data():
	msg4=None
	session['filename']=session['company']+'_'+str(uuid.uuid1())
	OP_RES_FOLDER=session['filename']
	OP_RES_UPLOAD=OP_RES_FOLDER+'/op_res_cats_files/'
	OP_RES_REPORT=OP_RES_FOLDER +'/Report/'
	
	# Get the name of the uploaded files
	uploaded_files = request.files.getlist("file[]")
	for file in uploaded_files:
		# Check if the file is one of the allowed types/extensions
		if file and allowed_file(file.filename):
			# Make the filename safe, remove unsupported chars
			filename = secure_filename(file.filename)
			if not os.path.exists(OP_RES_FOLDER):
				os.makedirs(OP_RES_FOLDER)
				os.makedirs(OP_RES_UPLOAD)
				os.makedirs(OP_RES_REPORT)
			# Move the file form the temporal folder to the upload
			
			file.save(os.path.join(OP_RES_UPLOAD, filename))
			filenames=os.listdir(OP_RES_UPLOAD)
			msg4=filenames
		else:
			msg4='Please select a valid extension (.xls or .xlsx)'
	return render_template('index_NOAM_company.html',msg4=msg4)



@application.route('/op_res_cats_upload', methods=['GET'])
@login_required
def op_res_cats_data_upload():
	msg=None
	msg2=None
	msg3=None
	msg_res=None
	msg4=None
	msg5=None
	msg6=None
	msg_ops=None
	OP_RES_FOLDER=session['filename']
	OP_RES_UPLOAD=OP_RES_FOLDER+'/op_res_cats_files/'
	OP_RES_REPORT=OP_RES_FOLDER +'/Report/'
	# Get the name of the uploaded files
	if len(os.listdir(OP_RES_UPLOAD))>0:
		process_res_cats.op_res_cats_files(path=OP_RES_UPLOAD,company=OP_RES_UPLOAD.split('_')[0],op_res_cats_report=OP_RES_REPORT)
		op_res_filenames=[f for f in os.listdir(OP_RES_REPORT) if f.endswith(('.xlsm','xlsx'))]
		content_res=''
		content_ops=''
		msg_res="RESOLUTION CATEGORIES VALIDATION:"
		if 'res_issues.txt' in os.listdir(OP_RES_REPORT):
			text_res=open(OP_RES_REPORT+'res_issues.txt', 'r+',encoding='utf8')
			content_res = text_res.read()
			text_res.close()
			if len(content_res)>1:
				msg='Errors found in Resolution Categories. Please check the Catalogue.'
			else:
				msg2='Your Resolution Categories are correct and ready to be uploaded.'
		else:
			msg3='Resolution Categories not found.'
		
		msg_ops="OPERATIONAL CATEGORIES VALIDATION:"
		if 'op_issues.txt' in os.listdir(OP_RES_REPORT):
			text_ops=open(OP_RES_REPORT+'op_issues.txt', 'r+',encoding='utf8')
			content_ops = text_ops.read()
			text_ops.close()
			if len(content_ops)>1:
				msg4='Errors found in Operational Categories. Please check the Catalogue.'
			else:
				msg5='Your Operational Categories are correct and ready to be uploaded.'
		else:
			msg6='Operational Categories not found.'
	return render_template('res_cats_upload.html', 
		op_res_filenames=op_res_filenames,
		text_res=content_res,
		text_ops=content_ops,
		msg_res=msg_res,
		msg_ops=msg_ops,
		msg=msg,msg2=msg2,msg3=msg3,msg4=msg4,msg5=msg5,msg6=msg6)


@application.route('/OP_RES_UPLOAD_Report/<filename>')
@login_required
def uploaded_RES_CATS_file(filename):
	OP_RES_FOLDER=session['filename']
	OP_RES_REPORT=OP_RES_FOLDER +'/Report/'
	return send_from_directory(OP_RES_REPORT,filename)

#############################################################
######ZTE split files
@application.route('/eia', methods=['GET','POST'])
@login_required
def eia_data():
	msg=None
	session['filename']=str(uuid.uuid1())
	EIA_FOLDER=session['filename']
	EIA_UPLOAD=EIA_FOLDER+'/eia_files/'
	EIA_REPORT=EIA_FOLDER +'/Report/'
	
	# Get the name of the uploaded files
	uploaded_files = request.files.getlist("file[]")
	for file in uploaded_files:
		# Check if the file is one of the allowed types/extensions
		if file and allowed_file(file.filename):
			# Make the filename safe, remove unsupported chars
			filename = secure_filename(file.filename)
			if not os.path.exists(EIA_FOLDER):
				os.makedirs(EIA_FOLDER)
				os.makedirs(EIA_UPLOAD)
				os.makedirs(EIA_REPORT)
			# Move the file form the temporal folder to the upload
			
			file.save(os.path.join(EIA_UPLOAD, filename))
			filenames=os.listdir(EIA_UPLOAD)
			msg=filenames
		else:
			msg='Please select a valid extension (.xls, .xlsx or .csv)'
	return render_template('index_eia.html',msg=msg)

@application.route('/eia_upload', methods=['GET'])
@login_required
def eia_upload():
	EIA_FOLDER=session['filename']
	EIA_UPLOAD=EIA_FOLDER+'/eia_files/'
	EIA_REPORT=EIA_FOLDER +'/Report/'
	# Get the name of the uploaded files
	if len(os.listdir(EIA_UPLOAD))>0:
		process_zte.split_zte_file(path=EIA_UPLOAD,report=EIA_REPORT)
		eia_filenames=os.listdir(EIA_REPORT)
	return render_template('eia_upload.html', eia_filenames=eia_filenames)



@application.route('/EIA_Report/<filename>')
@login_required
def uploaded_EIA_file(filename):
	EIA_FOLDER=session['filename']
	EIA_REPORT=EIA_FOLDER +'/Report/'
	return send_from_directory(EIA_REPORT,filename)




@application.route('/logout', methods=['GET'])
@login_required
def logout():
	ID_FOLDER=session['filename']
	if os.path.exists(ID_FOLDER):
		shutil.rmtree(ID_FOLDER)
	logout_user()
	return redirect('login')


# create login manager
login_manager = LoginManager()
# init login manager on application
login_manager.init_app(application)
# set login view
login_manager.login_view = 'login_get'
# callback to reload the user object
@login_manager.user_loader
def load_user(userid):
	return USERS[userid]
if __name__=='__main__':
	application.run(debug=True)