import json
from typing import List

from pandas import DataFrame
import pandas as pd

from config.globals import DevelopmentConfig

data_config = DevelopmentConfig()


def column_values(df: DataFrame, column_name):
    values = list(set(df.groupby(column_name).size().reset_index()[column_name]))
    values.sort()
    return values


class ConfigPack(object):
    def __init__(self, name, values=None, selected=None):
        if selected is None:
            selected = []
        if values is None:
            values = []
        self.name, self.values, self.selected = name, values, selected


class ComputeMethod(object):
    def __init__(self, method_name, name, values, selected=None):
        if selected is None:
            selected = []
        self.method_name = method_name
        self.name = name
        self.values = values
        self.selected = selected


method_divided_by = ComputeMethod('divided_by', 'Divided By', ['Total', 'ICE', 'BEV', 'PHEV'])


class GlobalConfig(object):
    def __init__(self, filters=None, page_include=ConfigPack('Slide Include...'),
                 pr_state=ConfigPack('PR_State'), pr_state_2=ConfigPack('PR_State')):
        if filters is None:
            filters = []
        self.name = 'global'
        self.filters = filters
        self.page_include = page_include
        self.pr_state = pr_state
        self.pr_state_2 = pr_state_2


class ComputeMethod(object):
    divided_by = ConfigPack("Divided By", ["Total", "ICE", "BEV", "PHEV"], [])


class SlideConfig(object):
    def __init__(self, name='', filters=None, group_by=ConfigPack('group_by'),
                 computer_methods=None, img='', included=False):
        if computer_methods is None:
            computer_methods = []
        if filters is None:
            filters = []
        self.name = name
        self.filters = filters
        self.group_by = group_by
        self.computer_methods = computer_methods
        self.img = img
        self.included = included


class ReportConfig(object):
    def __init__(self, global_config: GlobalConfig, slides_config: List[SlideConfig]):
        self.Global = global_config
        self.Slides = slides_config


def gen_config(df: DataFrame):
    pr_status = ConfigPack(name='OEM', values=column_values(df, 'pr_status'))
    oem = ConfigPack(name='OEM', values=column_values(df, 'oem'))
    brand = ConfigPack(name='Brand', values=column_values(df, 'brand'))
    build_type = ConfigPack(name='Build_Type', values=column_values(df, 'build_type'))
    year = ConfigPack(name="YEAR", values=column_values(df, 'year'))
    fuel_type = ConfigPack(name="Fuel_Type", values=column_values(df, 'fuel_type'))
    fuel_type_group = ConfigPack(name="Fuel_Type_Group", values=column_values(df, 'fuel_type_group'))

    global_config = GlobalConfig(filters=[oem, brand, build_type, year], pr_state=pr_status, pr_state_2=pr_status)

    slide1 = SlideConfig('page 1', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p1.png', True)
    slide2 = SlideConfig('page 2', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p2.png', True)
    slide3 = SlideConfig('page 3', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p3.png', True)
    slide4 = SlideConfig('page 4', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p4.png', True)
    slide5 = SlideConfig('page 5', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p5.png', True)
    slide6 = SlideConfig('page 6', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p6.png', True)
    slide7 = SlideConfig('page 7', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p7.png', True)
    slide8 = SlideConfig('page 8', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p8.png', True)
    slide9 = SlideConfig('page 9', [fuel_type], ConfigPack('Group By', ['Brand', 'Oem']),
                         'p9.png', True)
    config_pack = ReportConfig(global_config, [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9])
    return config_pack


def json2obj(json_data):
    configs = ReportConfig(**json.loads(json_data))
    return configs


if __name__ == '__main__':
    df = pd.read_csv('../sample/planning.csv')
    df.columns = df.columns.map(lambda x: x.lower().strip().replace(' ', '_').replace('/', '_'))
    print(json.dumps(gen_config(df), default=lambda o: o.__dict__, indent=2))
