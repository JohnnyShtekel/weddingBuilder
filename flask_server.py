# -*- coding: utf-8 -*-
from flask import Flask, send_file, request
from werkzeug.exceptions import BadRequestKeyError
import json
from flask import Flask, send_file, request
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import werkzeug

app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = set(['xls', 'xlsx'])

@app.route('/')
def multiple_routes(**kwargs):
    return send_file('templates/index.html')


@app.route('/upload', methods=['POST'])
def get_xl_file():
     file = request.files['file']
     print file
     return json.dumps({'success':True}), 200, {'ContentType':'application/json'}


if __name__ == '__main__':
    app.debug = True
    app.run(host='127.0.0.1', debug=True)
