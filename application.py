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
application.config['CMDB_FOLDER'] = CMDB_FOLDER


ID_FOLDER=str(uuid.uuid1())
#COMPANY_FOLDER=''
#UPLOAD_FOLDER= ''
#DOWNLOAD_FOLDER=''
#ITSM_FOLDER=''
ITSM_FOLDER=ID_FOLDER + '/ITSM_sites'
UPLOAD_FOLDER=ID_FOLDER + '/File_to_validate'
#DOWNLOAD_FOLDER=id_folder + '/Report'
#application.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
#application.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
#application.config['ITSM_FOLDER'] = ITSM_FOLDER
#
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
    #msg= None
    #if request.method == 'POST':
    #    company = request.form['company']
    #    msg = 'Successfull'
    return render_template('home.html')#,msg=msg)


@application.route('/files', methods=['GET','POST'])
@login_required
def sites_history():
    #global COMPANY_FOLDER
    global ID_FOLDER
    global ITSM_FOLDER
    #global DOWNLOAD_FOLDER
    msg=None
    #msg2=None
    if request.method == 'POST':
        #company = request.form['company']
        #COMPANY_FOLDER=company
        #ID_FOLDER=COMPANY_FOLDER + '_' + ID_FOLDER
        #msg2 = 'Successfull'
        if 'file' not in request.files:
            print('No file attached in request')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            os.makedirs(ID_FOLDER)
            os.makedirs(ITSM_FOLDER)
            
            
            #os.makedirs(DOWNLOAD_FOLDER)
            file.save(os.path.join(ITSM_FOLDER, filename))
            msg=filename
        else:
            msg='Please select a valid extension (.xls or .xlsx)'
    return render_template('multi_upload_index.html',msg=msg)#,ms2=msg2)
# --- login manager ------------------------------------------------------------


@application.route('/upload', methods=['POST'])
def upload():
	global UPLOAD_FOLDER
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


