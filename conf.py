import os
import time

"""
1、python中都是从0开始计数的

2、source_table:源数据表
3、source_sheet:源数据表中的sheet名字
4、title_rowx:源数据表中标题所在的行数
5、filter_date:源数据表中使用此参数进行日期过滤
6、filter_department:源数据表中使用此参数进行部门过滤
7、target_table_path:目标保存数据表的目录
8、target_sheet:目标数据保存sheet
"""
# 数据源工作簿
source_table = r"E:\test\bb\A表-收料记录.xlsx"
# 数据源工作表
source_sheet = "Sheet1"
# 目标工作簿
target_table_path = r"E:\test\bb\4月境外料汇总.xlsx"
# 过滤部门
departments = [16, 19]
# 过滤时间
# filter_date = time.strftime("%Y/%m/%d")
filter_date = "2019/04/25"