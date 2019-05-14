# -*- coding: utf-8 -*-

from flask import Flask
from flask import redirect
from flask import render_template
from flask import request
from flask import flash
from flask_login import LoginManager
from flask_login import UserMixin # subclass of flask user
from flask_login import login_required
from flask_login import login_user
from flask_login import logout_user
import uuid
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
application.secret_key = str(uuid.uuid1())

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
        flash(application.secret_key)
        return render_template('login.html')


@application.route('/logout', methods=['GET'])
@login_required
def logout():
    logout_user()
    return redirect('login')



# home route

@application.route('/home', methods=['GET'])
@login_required
def home():
    msg=application.secret_key
    return render_template('home.html',msg=msg)


# --- login manager ------------------------------------------------------------

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
        port=5000, 
        threaded=True
    )


