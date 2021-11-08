from datetime import datetime
def getTodayYearMonthDayHourMinSec():
    "返回格式为：'20210503_02_11_13'"
    return datetime.today().strftime('%Y%m%d_%H_%M_%S')
