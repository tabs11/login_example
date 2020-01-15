# -*- coding: utf-8 -*-
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
import process_zte
import update_priority as up_prio
#from flask_caching import Cache


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

#config = {
#    "DEBUG": True,          # some Flask specific configs
#    "CACHE_TYPE": "simple", # Flask-Caching related configs
#    "CACHE_DEFAULT_TIMEOUT": 300
#}
#
## application base
application = Flask(__name__)
# tell Flask to use the above defined config
#application.config.from_mapping(config)
#cache = Cache(application)

SECRET_KEY='bla'#str(uuid.uuid1())
application.secret_key = SECRET_KEY
CMDB_FOLDER = 'CMDB_templates/'
application.config['CMDB_FOLDER']=CMDB_FOLDER

# These are the extension that we are accepting to be uploaded
application.config['ALLOWED_EXTENSIONS'] = set(['xlsx','xls','csv'])
# default route
@application.route('/', methods=['GET'])
def index():
	
	return redirect("/home", code=302)

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
		#session['username']=str(uuid.uuid1())
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



# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
	return '.' in filename and \
		   filename.rsplit('.', 1)[1] in application.config['ALLOWED_EXTENSIONS']



@application.route('/return-file/')
@login_required
def return_file():
	filename='Templates.zip'
	return send_file(os.path.join(application.config['CMDB_FOLDER'])+filename,attachment_filename=filename, as_attachment=True)


@application.route('/home')
@login_required
def file_downloads():
	return render_template('home.html')


# home route

@application.route('/home', methods=['POST'])
@login_required
def home():
	data=[s for s in os.listdir(os.getcwd()) if len(s) > 20]
	paths_to_del=[]
	dates=[]
	for i in range(len(data)):
		paths_to_del.append(os.getcwd()+ '/' + data[i])
		dates.append((dt.datetime.now()-datetime.fromtimestamp(os.path.getctime(paths_to_del[i]))).days)
		if dates[i]>0:
			shutil.rmtree(paths_to_del[i])	
		else:
			None
	msg=None
	if request.method == 'POST':
		company = request.form['company']
		session['company']=company
		msg = 'Successfull'
			#msg=dates
	return render_template('home.html',msg=msg)

#######
@application.route('/cmdb', methods=['GET','POST'])
@login_required
def cmdb():
	return render_template('cmdb_validation.html')
#######

@application.route('/files', methods=['GET','POST'])
@login_required
def sites_history():
	session['filename']=session['company']+'_'+str(uuid.uuid1())
	ID_FOLDER=session['filename']
	ITSM_FOLDER=ID_FOLDER + '/ITSM_sites/'
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
	msg=None
	if request.method == 'POST':
		if 'file' not in request.files:
			print('No file attached in request')
			return redirect(request.url)
		file = request.files['file']
		if file.filename == '':
			print('No file selected')
			return redirect(request.url)
		if file and allowed_file(file.filename):
			os.makedirs(ID_FOLDER)
			os.makedirs(ITSM_FOLDER)
			os.makedirs(UPLOAD_FOLDER)
			filename = secure_filename(file.filename)
			file.save(os.path.join(ITSM_FOLDER, filename))
			msg=filename
		else:
			msg='Please select a valid extension (.xls or .xlsx)'
	return render_template('multi_upload_index.html',msg=msg)


@application.route('/data', methods=['GET','POST'])
@login_required
def data_to_validate():
	ID_FOLDER=session['filename']
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	msg3=None
	uploaded_files = request.files.getlist("file[]")
	for file in uploaded_files:
		# Check if the file is one of the allowed types/extensions
		if file and allowed_file(file.filename):
			# Make the filename safe, remove unsupported chars
			filename = secure_filename(file.filename)
			if not os.path.exists(UPLOAD_FOLDER):
				os.makedirs(UPLOAD_FOLDER)
			
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
	msg=None
	msg2=None
	msg3=None
	msg4=None
	msg5=None
	msg6=None
	ID_FOLDER=session['filename']
	ITSM_FOLDER=ID_FOLDER + '/ITSM_sites/'
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	os.makedirs(DOWNLOAD_FOLDER)
	if len(os.listdir(UPLOAD_FOLDER))>0:
		process_data.process_file(path=UPLOAD_FOLDER,company=ID_FOLDER.split('_')[0],report=DOWNLOAD_FOLDER,history=ITSM_FOLDER)
		filenames = [f for f in os.listdir(DOWNLOAD_FOLDER) if f.endswith('.xlsx')] 
		if 'errorsCIs.txt' in os.listdir(DOWNLOAD_FOLDER):
			text_errors_Cis=open(DOWNLOAD_FOLDER+'errorsCIs.txt', 'r+',encoding='utf8')
			content_errors_Cis = text_errors_Cis.read()
			text_errors_Cis.close()
			if len(content_errors_Cis)>1:
				msg='YOUR CIS DATA HAS ISSUES. PLEASE DOWNLOAD THE REPORT AND CHECK THE SHEETS'
			else:
				content_errors_Cis=''
				msg2='YOUR CIS DATA HAS NO ERRORS. IS READY TO BE UPLOADED. CHECK THE POSSIBLE WARNINGS'
		else:
			content_errors_Cis=''
			msg3='CIS FILE NOT FOUND'


			
			
		if 'errorsSites.txt' in os.listdir(DOWNLOAD_FOLDER):
			text_errors_Sites=open(DOWNLOAD_FOLDER+'errorsSites.txt', 'r+',encoding='utf8')
			content_errors_Sites = text_errors_Sites.read()
			text_errors_Sites.close()
			if len(content_errors_Sites)>1:
				msg4='YOUR SITES DATA HAS ISSUES. PLEASE DOWNLOAD THE REPORT AND CHECK THE SHEETS'
			else:
				content_errors_Sites=''
				msg5='YOUR SITES DATA HAS NO ERRORS. IS READY TO BE UPLOADED. CHECK THE POSSIBLE WARNINGS'
		else:
			content_errors_Sites=''
			msg6='SITES FILE NOT FOUND'

		

		#text_correct_Sites=open(DOWNLOAD_FOLDER+'correct_dataSites.txt', 'r+',encoding='utf8')
		#content_correct_Sites = text_correct_Sites.read()
		#text_correct_Sites.close()
		#text_errors_report_Sites=open(DOWNLOAD_FOLDER+'errors_reportSites.txt', 'r+',encoding='utf8')
		#content_errors_report_Sites = text_errors_report_Sites.read()
		#text_errors_report_Sites.close()
#
		#text_correct_Cis=open(DOWNLOAD_FOLDER+'correct_dataCIs.txt', 'r+',encoding='utf8')
		#content_correct_Cis = text_correct_Cis.read()
		#text_correct_Cis.close()
		#text_errors_report_Cis=open(DOWNLOAD_FOLDER+'errors_reportCIs.txt', 'r+',encoding='utf8')
		#content_errors_report_Cis = text_errors_report_Cis.read()
		#text_errors_report_Cis.close()

		
		text_warnings=open(DOWNLOAD_FOLDER+'warnings.txt', 'r+',encoding='utf8')
		content_warnings = text_warnings.read()
		text_warnings.close()
		text_summary=open(DOWNLOAD_FOLDER+'summary.txt', 'r+',encoding='utf8')
		content_summary = text_summary.read()
		text_summary.close()
		

	return render_template('multi_files_upload.html', 
		filenames=filenames,
		text_errors_Cis=content_errors_Cis,
		text_errors_Sites=content_errors_Sites,
		#text_correct_Sites=content_correct_Sites,
		#text_correct_Cis=content_correct_Cis,
		#text_errors_report_Sites=content_errors_report_Sites,
		#text_errors_report_Cis=content_errors_report_Cis,
		text_warnings=content_warnings,
		text_summary=content_summary,msg=msg,msg2=msg2,msg3=msg3,msg4=msg4,msg5=msg5,msg6=msg6)


@application.route('/report/<filename>')
@login_required
def uploaded_file(filename):
	ID_FOLDER=session['filename']
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	return send_from_directory(DOWNLOAD_FOLDER,filename)



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
	OP_RES_FOLDER=session['filename']
	OP_RES_UPLOAD=OP_RES_FOLDER+'/op_res_cats_files/'
	OP_RES_REPORT=OP_RES_FOLDER +'/Report/'
	# Get the name of the uploaded files
	if len(os.listdir(OP_RES_UPLOAD))>0:
		process_res_cats.op_res_cats_files(file_path=OP_RES_UPLOAD,company=OP_RES_UPLOAD.split('_')[0],op_res_cats_report=OP_RES_REPORT)
		op_res_filenames=os.listdir(OP_RES_REPORT)
		text = open(OP_RES_REPORT+'issues.txt', 'r+',encoding='utf8')
		content = text.read()
		text.close()
	return render_template('res_cats_upload.html', op_res_filenames=op_res_filenames,text=content)


@application.route('/OP_RES_UPLOAD_Report/<filename>')
@login_required
def uploaded_RES_CATS_file(filename):
	OP_RES_FOLDER=session['filename']
	OP_RES_REPORT=OP_RES_FOLDER +'/Report/'
	return send_from_directory(OP_RES_REPORT,filename)
#############################################################

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

#port = int(os.environ.get("PORT", 5000))

if __name__ == '__main__':
	#application.run(
	#	host='0.0.0.0', 
	#	port=port
	#)
	application.run(
		#host='0.0.0.0', 
		#port=3000, 
		threaded=True
	)
