from datetime import datetime
def getTodayYearMonthDayHourMinSec():
    "返回格式为：'20210503_02_11_13'"
    return datetime.today().strftime('%Y%m%d_%H_%M_%S')

def getDay():
    "返回格式2022年1月18日"
    return datetime.today().strftime('%Y年%m月%d日')