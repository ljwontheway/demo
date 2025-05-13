import datetime

def is_holiday(date: datetime.date, holidays=None, extra_workdays=None, extra_holidays=None):
    """
    判断某个日期是否为节假日（含法定节假日、周末、调休）
    :param date: datetime.date 对象
    :param holidays: 法定节假日列表，元素为 'YYYY-MM-DD' 字符串
    :param extra_workdays: 调休上班日列表，元素为 'YYYY-MM-DD' 字符串
    :param extra_holidays: 调休放假日列表，元素为 'YYYY-MM-DD' 字符串
    :return: True/False
    """
    # 2025年法定节假日（以国务院公告为准，以下为常规推算，具体以官方为准）
    default_holidays = [
        '2025-01-01', # 元旦
        '2025-01-29', '2025-01-30', '2025-01-31', '2025-02-01', '2025-02-02', '2025-02-03', '2025-02-04', # 春节
        '2025-04-04', '2025-04-05', '2025-04-06', # 清明节
        '2025-05-01', '2025-05-02', '2025-05-03', '2025-05-04', '2025-05-05', # 劳动节
        '2025-05-31', '2025-06-01', '2025-06-02', # 端午节
        '2025-09-05', '2025-09-06', '2025-09-07', # 中秋节
        '2025-10-01', '2025-10-02', '2025-10-03', '2025-10-04', '2025-10-05', '2025-10-06', '2025-10-07', # 国庆节
    ]
    # 2025年部分调休上班日（以国务院公告为准，以下为常规推算，具体以官方为准）
    default_extra_workdays = [
        '2025-01-26', # 春节调休上班（周日）
        '2025-02-08', # 春节调休上班（周六）
        '2025-04-07', # 清明调休上班（周一）
        '2025-04-27', # 劳动节调休上班（周六）
        '2025-09-13', # 中秋调休上班（周六）
        '2025-10-11', # 国庆调休上班（周六）
    ]
    default_extra_holidays = [
        # 如有特殊调休放假日可补充
    ]

    holidays = holidays or default_holidays
    extra_workdays = extra_workdays or default_extra_workdays
    extra_holidays = extra_holidays or default_extra_holidays

    date_str = date.strftime('%Y-%m-%d')

    # 优先判断调休
    if date_str in extra_workdays:
        return False  # 调休上班日
    if date_str in extra_holidays:
        return True   # 调休放假日

    # 法定节假日
    if date_str in holidays:
        return True

    # 周末（周六=5，周日=6）
    if date.weekday() >= 5:
        return True

    return False

# 测试用例
if __name__ == "__main__":
    test_dates = [
        '2025-01-01', # 元旦
        '2025-01-31', # 春节
        '2025-03-01', # 普通工作日
        '2025-05-01', # 劳动节
        '2025-07-01', # 普通工作日
        '2025-10-03', # 国庆节
        '2025-01-26', # 春节调休上班（周日，应为工作日）
        '2025-02-08', # 春节调休上班（周六，应为工作日）
        '2025-06-07', # 普通周六，应为休息日
        '2025-06-02', # 端午节
    ]
    for d in test_dates:
        dt = datetime.datetime.strptime(d, '%Y-%m-%d').date()
        print(f"{d} 是否节假日: {is_holiday(dt)}") 