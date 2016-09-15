# -*- coding: utf-8 -*-
import json
import os
import datetime
from flask import Flask, request
from flask import send_file
from werkzeug.utils import secure_filename
from file_mangaer import FileManager
from gvia_daily_report_for_chosen_month import GviaDailyPerMonth
from team_gvia_for_chosen_month import TeamGviaForMonth

app = Flask(__name__, static_folder='static_gvia_yadim')
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xls', 'xlsx'}


@app.route('/api/v1/gvia-yadim-report/')
def multiple_routes(**kwargs):
    return send_file('templates/index.html')



@app.route('/api/v1/gvia-yadim-report/download_report/<file_name>/')
def download_report(file_name):
    try:
        return send_file(file_name)
    except:
        return json.dumps({'error': True}), 500, {'ContentType': 'application/json'}



@app.route('/api/v1/gvia-yadim-report/upload/', methods=['POST'])
def get_xl_file():
    try:
        chosen_file = request.files['file']
        worker_name = request.form['worker']
        manager_name = request.form['manager']
        if chosen_file and allowed_file(chosen_file.filename):
            filename = secure_filename(chosen_file.filename)
            chosen_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            fn = FileManager()
            fn.extract_all_comments_and_update_cem(os.path.join(app.config['UPLOAD_FOLDER'], filename), worker_name, manager_name)
            return json.dumps({'success': True}), 200, {'ContentType': 'application/json'}
        else:
            return json.dumps({'error': True}), 500, {'ContentType': 'application/json'}
    except Exception as e:
        return json.dumps({'error': str(e)}), 500, {'ContentType': 'application/json'}




@app.route('/api/v1/gvia-yadim-report/run_department_report/', methods=['POST'])
def run_department_report():
    # try:
        year = int(request.form['year'])
        month = int(request.form['month'])
        gvia_daily = GviaDailyPerMonth('report_for_hani', requested_year=year, requested_month=month)
        gvia_daily.run_report()
        gvia_team = TeamGviaForMonth(requested_month=month, requested_year=year)
        report_file_name = gvia_team.run_report()
        return json.dumps({'file_name': report_file_name}), 200, {'ContentType': 'application/json'}
    # except Exception as e:
    #     return json.dumps({'error': str(e)}, 500, {'ContentType': 'application/json'})



def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']


if __name__ == '__main__':
    app.debug = True
    # app.run(host='0.0.0.0', debug=True, port=5000)
    app.run(host='127.0.0.1', debug=True, port=8080)
