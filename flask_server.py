# -*- coding: utf-8 -*-
import os
import json
from flask import Flask, send_file, request
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
from file_mangaer import FileManager

app = Flask(__name__, static_folder='static_gvia_yadim')
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = set(['xls', 'xlsx'])

@app.route('/api/v1/gvia_yadim_report/')
def multiple_routes(**kwargs):
    return send_file('templates/index.html')


@app.route('/api/v1/gvia_yadim_report/upload/', methods=['POST'])
def get_xl_file():
    try:

         file = request.files['file']
         workerName = request.form['worker']
         if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                fn = FileManager()
                fn.extract_all_comments_and_update_cem(os.path.join(app.config['UPLOAD_FOLDER'], filename),workerName)
                return json.dumps({'success':True}), 200, {'ContentType':'application/json'}
         else:
             return json.dumps({'error':False}), 200, {'ContentType':'application/json'}


    except:
            return json.dumps({'error':False}), 200, {'ContentType':'application/json'}



def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']


if __name__ == '__main__':
    app.debug = True
    app.run(host='0.0.0.0', debug=True, port=5001)
