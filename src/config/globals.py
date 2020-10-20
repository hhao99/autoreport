from os.path import dirname, join


class WebConfig(object):
    DEBUG = False
    TESTING = False
    CSRF_ENABLE = True
    DATA_PATH = './data.xlsx'
    PLANNING_DATA = './planning.csv'


class DevelopmentConfig(WebConfig):
    DEBUG = True
    DEVELOPMENT = True
    DATA_PATH = join(dirname(__file__), "../sample/data.xlsx")
    PLANNING_DATA = join(dirname(__file__), "../sample/planning.csv")


class ProductionConfig(WebConfig):
    DEBUG = False
