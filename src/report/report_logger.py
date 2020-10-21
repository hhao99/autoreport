import logging

logging.basicConfig(filename='c:/auto-report/mylog.txt', \
                    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)


# logging.warning('warn message')
# logging.info('info message')
# logging.debug('debug message')
# logging.error('error message')
# logging.critical('critical message')

def record_log(report_id, vw_or_mkt, pr_status, year, fuel_type='', fuel_type_group='', brand_or_oem=''):
    info = "Report=" + str(report_id) + ",DataSource=" + str(vw_or_mkt) + ",PR_status=" + str(pr_status) + ", year=" \
           + str(year) + ",fuel_type=" + str(fuel_type) + ",fuel_type_group=" + str(fuel_type_group) \
           + ",brand_or_oem = " + str(brand_or_oem) + " Data is missing!"
    logging.info(info)
