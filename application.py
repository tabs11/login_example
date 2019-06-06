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
	'jschuur' : User('jschuur','teste')
	#'user3' : User('user3','cenas')
}

# application base
application = Flask(__name__)
SECRET_KEY='bla'#str(uuid.uuid1())
application.secret_key = SECRET_KEY
CMDB_FOLDER = 'CMDB_templates/'
application.config['CMDB_FOLDER']=CMDB_FOLDER

# These are the extension that we are accepting to be uploaded
application.config['ALLOWED_EXTENSIONS'] = set(['xlsx','xls'])
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
		flash("username ou password erradas")
		return render_template('login.html')



# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
	return '.' in filename and \
		   filename.rsplit('.', 1)[1] in application.config['ALLOWED_EXTENSIONS']



@application.route('/return-file/')
@login_required
def return_file():
	filename='cmdb_templates.zip'
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
		dates.append((dt.datetime.now()-datetime.fromtimestamp(os.path.getctime(paths_to_del[i]))).seconds)
		if dates[i]>60*60*24:
			shutil.rmtree(paths_to_del[i])
		else:
			None
	msg=None
	if request.method == 'POST':
		company = request.form['company']
		session['company']=company
		msg = 'Successfull'
	return render_template('home.html',msg=msg)

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
@login_required
def upload():
	ID_FOLDER=session['filename']
	ITSM_FOLDER=ID_FOLDER + '/ITSM_sites/'
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate/'
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	os.makedirs(DOWNLOAD_FOLDER)
	if len(os.listdir(UPLOAD_FOLDER))>0:
		process_data.process_file(path=UPLOAD_FOLDER,company=ID_FOLDER.split('_')[0],report=DOWNLOAD_FOLDER,history=ITSM_FOLDER)
		filenames=os.listdir(DOWNLOAD_FOLDER)
		text = open(DOWNLOAD_FOLDER+'issues.txt', 'r+',encoding='utf8')
		content = text.read()
		text.close()
	return render_template('multi_files_upload.html', filenames=filenames,text=content)


@application.route('/report/<filename>')
@login_required
def uploaded_file(filename):
	ID_FOLDER=session['filename']
	DOWNLOAD_FOLDER=ID_FOLDER+'/Report/'
	return send_from_directory(DOWNLOAD_FOLDER,filename)



@application.route('/noam', methods=['GET', 'POST'])
def comp_noam():
	msg= None
	if request.method == 'POST':
		company_noam = request.form['company_noam']
		session['company_noam']=company_noam
		
		msg = 'Successfull'
	return render_template('index_NOAM_company.html',msg=msg)


@application.route('/noam_data', methods=['GET','POST'])
def noam_data():
	msg3=None
	session['filename_final']=session['company_noam']+'_'+str(uuid.uuid1())
	NOAM_FOLDER=session['filename_final']
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
	NOAM_FOLDER=session['filename_final']
	NOAM_UPLOAD=NOAM_FOLDER+'/NOAM_files/'
	NOAM_REPORT=NOAM_FOLDER +'/Report/'
	# Get the name of the uploaded files
	if len(os.listdir(NOAM_UPLOAD))>0:
		process_noam_data.noam_files(file_path=NOAM_UPLOAD,company=NOAM_FOLDER.split('_')[0],NOAM_report=NOAM_REPORT)
		noam_filenames=os.listdir(NOAM_REPORT)

	return render_template('noam_files_upload.html', noam_filenames=noam_filenames)





@application.route('/NOAM_Report/<filename>')
def uploaded_NOAM_file(filename):
	NOAM_FOLDER=session['filename_final']
	NOAM_REPORT=NOAM_FOLDER +'/Report/'
	return send_from_directory(NOAM_REPORT,filename)




@application.route('/logout', methods=['GET'])
@login_required
def logout():
	#ID_FOLDER=session['filename']
	#if os.path.exists(ID_FOLDER):
	#	shutil.rmtree(ID_FOLDER)
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


if __name__ == '__main__':
	application.run(
		host='0.0.0.0', 
		port=3000, 
		threaded=True
	)
