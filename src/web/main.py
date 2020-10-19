from flask import Flask, make_response, jsonify
from config.globals import DevelopmentConfig
import pandas as pd
import traceback

app = Flask(__name__)
app.config.from_object(DevelopmentConfig())


@app.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 'Not found'}), 404)


@app.errorhandler(500)
def server_error(error):
    return make_response(jsonify({'error': 'Internal Server Error:' + traceback.print_stack()}), 500)


@app.route('/')
def index():
    return "Report Automation Backend services - version 0.1 20200922"


@app.route('/get/<field>')
def field_query(field):
    df = pd.read_csv(app.config['PLANNING_DATA'])
    df.columns = df.columns.map(lambda x: x.lower().strip().replace(' ', '_').replace('/', '_'))
    values = df.groupby(field.lower()).size().reset_index()[field.lower()]
    return jsonify(list(set(values)))


@app.route('/queryby/<field>/<value>/<target_field>')
def query_by(field, value, target_field):
    df = pd.read_excel(app.config['DATA_PATH'])
    df1 = df[df[field] == value]
    return jsonify(list(set(df1[target_field])))



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app.run()
