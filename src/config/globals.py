from os.path import dirname, join


class GlobalConfig(object):
    DEBUG = False
    TESTING = False
    CSRF_ENABLE = True
    DATA_PATH = './data.xlsx'


class DevelopmentConfig(GlobalConfig):
    DEBUG = True
    DEVELOPMENT = True
    DATA_PATH = join(dirname(__file__), "../sample/data.xlsx")


class ProductionConfig(GlobalConfig):
    DEBUG = False
