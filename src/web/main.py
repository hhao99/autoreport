from flask import Flask, jsonify
from flask_restful import Api, Resource
from config import DevelopmentConfig, ProductionConfig

app = Flask(__name__)
app.config.from_object(DevelopmentConfig())

@app.route('/')
def index():
    return "Report Automation Backend services - version 0.1 20200922"

@app.route('/get/<field>')
def field_query(field):
    import pandas as pd
    from flask import jsonify

    try:
        df = pd.read_excel(app.config['DATA_PATH'])
        return jsonify((list(set(df[field]))))
    except:
        return 'field exception'

@app.route('/queryby/<field>/<value>/<target_field>')
def query_by(field,value,target_field):
    import pandas as pd
    from flask import jsonify
    try:
        df = pd.read_excel(app.config['DATA_PATH'])
        df1 = df[df[field] == value]
        return jsonify(list(set(df1[target_field])))
    except:
        return "exception"




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app.run()
