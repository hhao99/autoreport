import json
from typing import List


class ReportConfig(object):
    def __init__(self, pr_round, year_range, filters, group_by, compute_methods):
        self.pr_round = pr_round
        self.year_rang = year_range
        self.filters = filters
        self.group_by = group_by
        self.compute_methods = compute_methods


class PRRound(object):
    def __init__(self, name, values, selected):
        self.name = name
        self.values = values
        self.selected = selected


class YearRange(object):
    def __init__(self, name, values, selected):
        self.name = name
        self.values = values
        self.selected = selected


class ReportFilter(object):
    def __init__(self, name, value_type, values, selected):
        self.name = name
        self.value_type = value_type
        self.values = values
        self.selected = selected


class GroupBy(object):
    def __init__(self, name, values, selected):
        self.name = name
        self.values = values
        self.selected = selected


class ComputeMethod(object):
    def __init__(self, name, values, selected):
        self.name = name
        self.values = values
        self.selected = selected


class ReportConfigs(object):
    def __init__(self, reports: List[ReportConfig]):
        self.reports = reports


def json2obj(json_data):
    configs = ReportConfigs(**json.loads(json_data))
    return configs


if __name__ == "__main__":
    with open('./config.json') as f:
        data = f.read().replace('\n', '')
        obj = json2obj(data)
        print(json.dumps(obj, default=lambda o: o.__dict__, indent=2))
