import json

import pandas as pd
from pandas import DataFrame

from config.globals import DevelopmentConfig
from config.report import GlobalConfig, ReportConfig, ConfigPack

data_config = DevelopmentConfig()


def column_values(df: DataFrame, column_name):
    values = list(set(df.groupby(column_name).size().reset_index()[column_name]))
    values.sort()
    return values


def gen_config(df: DataFrame):
    pr_status = ConfigPack(name='OEM', values=column_values(df, 'pr_status'))
    oem = ConfigPack(name='OEM', values=column_values(df, 'oem'))
    brand = ConfigPack(name='Brand', values=column_values(df, 'brand'))
    build_type = ConfigPack(name='Build_Type', values=column_values(df, 'build_type'))
    year = ConfigPack(name="YEAR", values=column_values(df, 'year'))
    global_config = GlobalConfig(filters=[oem, brand, build_type, year], pr_state=pr_status, pr_state_2=pr_status)

    config_pack = ReportConfig(global_config, [])
    return config_pack


if __name__ == '__main__':
    df = pd.read_csv('../sample/planning.csv')
    df.columns = df.columns.map(lambda x: x.lower().strip().replace(' ', '_').replace('/', '_'))
    print(json.dumps(gen_config(df), default=lambda o: o.__dict__, indent=2))
