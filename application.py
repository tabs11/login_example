# -*- coding: utf-8 -*-

from flask import Flask, request, redirect, url_for, render_template, send_from_directory,send_file,flash
from werkzeug.utils import secure_filename
from flask_login import LoginManager
from flask_login import UserMixin # subclass of flask user
from flask_login import login_required
from flask_login import login_user
from flask_login import logout_user
import shutil
import uuid
import os
import datetime as dt 
from datetime import datetime
class User(UserMixin):
	def __init__(self, username, password):
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
	'user1' : User('user1','pass'),
	'user2' : User('user2','teste'),
	'user3' : User('user3','cenas')
}

# application base
application = Flask(__name__)
application.secret_key = 'bla'

CMDB_FOLDER = 'CMDB_templates/'
#application.config['CMDB_FOLDER']=CMDB_FOLDER
#
#ITSM_FOLDER=ID_FOLDER+'/ITSM_sites'
##application.config['ITSM_FOLDER'] = ID_FOLDER + ITSM_FOLDER
#UPLOAD_FOLDER=ID_FOLDER+'/Files_to_validate'
##application.config['UPLOAD_FOLDER'] = ID_FOLDER + UPLOAD_FOLDER
ID_FOLDER=''
ITSM_FOLDER=''
UPLOAD_FOLDER=''
#ID_FOLDER = application.secret_key
#ITSM_FOLDER=ID_FOLDER+'/ITSM_sites'
#UPLOAD_FOLDER=ID_FOLDER+'/Files_to_validate'


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
	global username
	# get details from post request
	username = request.form['username']
	password = request.form['password']
	# get  user
	try:
		user = USERS[username]
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


@application.route('/logout', methods=['GET'])
@login_required
def logout():
	logout_user()
	return redirect('login')


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
	return render_template('home.html')

@application.route('/files', methods=['GET','POST'])
@login_required
def sites_history():
	global ID_FOLDER
	global ITSM_FOLDER
	#global UPLOAD_FOLDER
	ID_FOLDER=str(uuid.uuid1())
	ITSM_FOLDER=ID_FOLDER + '/ITSM_sites'
	#UPLOAD_FOLDER=ID_FOLDER + '/File_to_validate'
	
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
			#os.makedirs(UPLOAD_FOLDER)
			filename = secure_filename(file.filename)
			file.save(os.path.join(ITSM_FOLDER, filename))
			msg=ID_FOLDER
		else:
			msg='Please select a valid extension (.xls or .xlsx)'
	return render_template('multi_upload_index.html',msg=msg)#,ms2=msg2)

@application.route('/upload', methods=['POST'])
@login_required
def upload():
	global ID_FOLDER
	global UPLOAD_FOLDER
	UPLOAD_FOLDER=ID_FOLDER + '/Files_to_validate'
	msg2=None
	# Get the name of the uploaded files
	uploaded_files = request.files.getlist("file[]")
	for file in uploaded_files:
		# Check if the file is one of the allowed types/extensions
		if file and allowed_file(file.filename):
			# Make the filename safe, remove unsupported chars
			filename = secure_filename(file.filename)
			os.makedirs(UPLOAD_FOLDER)
			# Move the file form the temporal folder to the upload
			file.save(os.path.join(UPLOAD_FOLDER, filename))
		else:
			msg2='Please select a valid extension (.xls or .xlsx)'
			return render_template('multi_upload_index.html',msg2=msg2)
	return render_template('multi_files_upload.html',msg2=msg2)


@application.route('/upload/<filename>')
def uploaded_file(filename):
	return send_from_directory(UPLOAD_FOLDER,filename)

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
