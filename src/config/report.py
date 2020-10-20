import json
from os.path import dirname, join
from typing import List


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


class SlideConfig(object):
    def __init__(self, name='', filters=None, group_by=ConfigPack('group_by'),
                 computer_methods: List[ComputeMethod] = None, img='', included=False):
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


def json2obj(json_data):
    configs = ReportConfig(**json.loads(json_data))
    return configs


if __name__ == "__main__":
    with open(join(dirname(__file__), '../sample/config.json')) as f:
        data = f.read().replace('\n', '')
        obj = json2obj(data)
        print(json.dumps(obj, default=lambda o: o.__dict__, indent=2))
