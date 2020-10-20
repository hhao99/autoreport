from report.report1 import Report1
from report.report2 import Report2


def RunAll(config):
    for i in range(2):
        report = CreatReport(i)
        report.LoadConfig('')
        report.LoadBaseData()
        report.FilterData()
        report.CreateSlide()
        report.Run()
        report.SaveSlide(i)
    print("test run all")


def CreatReport(reportid):
    if reportid == 0:
        report = Report1(reportid)
    elif reportid == 1:
        report = Report2(reportid)
    else:
        report = Report1(reportid)
    return report


if __name__ == '__main__':
    RunAll("")
