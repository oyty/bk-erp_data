import xlrd
from datetime import datetime
from datetime import date
from xlrd import xldate_as_tuple
import leancloud
import json

data = xlrd.open_workbook('bk-diesel-data.xlsx')
table = data.sheets()[0]
nrows = table.nrows

leancloud.init("akNEsr5TPApIehyBBMY4spgu-gzGzoHsz", "UwiO0Dd5BOrC2ySAugrDEqmC")

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""

    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    raise TypeError ("Type %s not serializable" % type(obj))

count = 0
for i in range(nrows):
    diesel_date = datetime(*xldate_as_tuple(table.row_values(i)[1], 0))
    oil_company = table.row_values(i)[2]
    work_team = table.row_values(i)[3]
    transport_cart_no = table.row_values(i)[4]
    diesel_type = table.row_values(i)[5]
    diesel_mount = table.row_values(i)[6]
    diesel_unit_price = table.row_values(i)[7]
    remark = table.row_values(i)[9]

    DieselData = leancloud.Object.extend('DieselData')
    todo_folder = DieselData()
    todo_folder.set('date', diesel_date)
    todo_folder.set('oil_company', oil_company)
    todo_folder.set('work_team', work_team)
    todo_folder.set('transport_cart_no', transport_cart_no)
    todo_folder.set('diesel_type', diesel_type)
    todo_folder.set('diesel_mount', diesel_mount)
    todo_folder.set('diesel_unit_price', diesel_unit_price)
    todo_folder.set('remark', remark)
    todo_folder.save()
    count += 1
    print('加入数据第 ' + str(count) + '条')

