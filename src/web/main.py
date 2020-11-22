import json

import traceback
from flask import Flask, jsonify, make_response
import pandas as pd
from config.page import gen_config
from config.globals import DevelopmentConfig
import report.report_launch as launch

app = Flask(__name__)
app.config.from_object(DevelopmentConfig())


@app.after_request
def cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Method'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'x-requested-with,content-type'
    return response


@app.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 'Not found'}), 404)


@app.errorhandler(500)
def server_error(error):
    return make_response(jsonify({'error': 'Internal Server Error:' + traceback.print_stack()}), 500)


@app.route('/')
def index():
    return "Report Automation Backend services - version 0.1 20200922"


@app.route('/get/<field>', methods=['GET'])
def field_query(field):
    df = pd.read_csv(app.config['PLANNING_DATA'])
    df.columns = df.columns.map(lambda x: x.lower().strip().replace(' ', '_').replace('/', '_'))
    values = df.groupby(field.lower()).size().reset_index()[field.lower()]
    return jsonify(list(set(values)))


@app.route('/queryby/<field>/<value>/<target_field>', methods=['GET'])
def query_by(field, value, target_field):
    df = pd.read_excel(app.config['DATA_PATH'])
    df1 = df[df[field] == value]
    return jsonify(list(set(df1[target_field])))


@app.route('/config', methods=['GET'])
def ppt_config():
    df = pd.read_csv(app.config['PLANNING_DATA'])
    df.columns = df.columns.map(lambda x: x.lower().strip().replace(' ', '_').replace('/', '_'))
    return json.dumps(gen_config(df), default=lambda o: o.__dict__, indent=2)


@app.route('/config', methods=['POST'])
def generate_report(config):
    launch.RunAll(None)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app.run()
