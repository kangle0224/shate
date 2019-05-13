import json
from codecs import open
from datetime import datetime, date, timedelta


def gen_date_list(num):
    """

    :param num: 若要生成当天的数据，num=1；若要输入从本月1号到今天的数据则输入今天的日期，如5
    :return:
    """
    date_list = []
    for i in range(num):
        date_list.append(str(date.today() - timedelta(days=i)).replace("-", "/"))
        date_list.reverse()
    print(date_list)
    return date_list

gen_date_list(13)
