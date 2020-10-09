import os

class Config(object):
    DEBUG = False
    TESTING = False
    CSRF_ENABLE = True
    DATA_PATH='./data.xlsx'

class DevelopmentConfig(Config):
    DEBUG = True
    DEVELOPMENT = True

class ProductionConfig(Config):
    DEBUG = False
