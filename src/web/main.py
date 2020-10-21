from flask import Flask, make_response, jsonify
from main.config import DevelopmentConfig
import pandas as pd
import traceback
from main.config import gen_config
import json

//need install the cross original support package
from flask_cors import CORS

def create_app():
    
    app = Flask(__name__)
    // enable the client js to request service without same orignal
    CORS(app)
    
    // todo: alter to the production before deployment
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


    @app.route('/config', methods=['GET','POST'])
    def data_config():
        df = pd.read_csv(app.config['PLANNING_DATA'])
        df.columns = df.columns.map(lambda x: x.lower().strip().replace(' ', '_').replace('/', '_'))
        return json.dumps(gen_config(df), default=lambda o: o.__dict__, indent=2)


    // update the config query and post methods, need the env CONFIG_FILE and CONFIG_DATA
    // CONFIG_FILE is the default app config
    // CONFIG_DATA is the customized config data
    @app.route('/report/config', methods=['POST'])
    def generate_report(config):
        if(request.method == 'GET'):
            with open(app.config['CONFIG_FILE']) as f:
                config = json.loads(f.read())
                return jsonify(config)
        else:
            if(request.is_json):
                print("got the json data")
                config = request.get_json()
                with open(app.config["CONFIG_DATA"], 'w') as f:
                    f.write(json.dumps(config))
                
            return "post success"


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    app = create_app()
    app.run()
