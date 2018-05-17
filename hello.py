# -*- coding:utf-8 -*-

"""
created by 2018-05-14
@author linwei

learn to operate the flask
"""

from flask import *

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    return '<h1>hello world!</h1>'

@app.route('/login', methods=['GET'])
def login_form():
    return '''<form action="/login" method="post">
                <p><input name="username"></p>
                <p><input name="password" type="password"></p>
                <p><button type="submit">Sign In</button></p>
                </form>'''

@app.route('/login', methods=['POST'])
def login():
    if request.form['username'] == 'linwei' and request.form['password'] == 'liulei':
        return '<h3>Hello, linwei</h3>'
    return '<h3>wrong username or password!</h3>'

@app.route('/data', methods=['GET'])
def get_json():
    list = ['linwei', 'liulei']
    return jsonify(list)

if __name__ == '__main__':
    app.run(debug=True)
