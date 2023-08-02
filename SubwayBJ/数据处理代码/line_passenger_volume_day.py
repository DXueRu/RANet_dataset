import datetime

import xlrd
from pymysql_lib import UsingMysql


def read_excel(path,date):
    # 打开文件
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    print(nrows, ncols)
    # 初始化参数
    data = []
    line_name = ""
    station_id = ""
    station_name = ""
    i = 6
    # 读取数据
    while i < nrows-1:
        row = table.row_values(i)
        row.append(date)
        print(row)
        data.append(row)
        # 判段哪条线
        # if "分票种进出站量日统计表" in str(row[0]):
        #     line_name = row[0][0:row[0].index("分票种进出站量日统计表")]
        #     i += 4
        #     continue
        # 去掉日合计和月累计
        # if row[0] == "小计":
        #     i += 6
        #     continue
        # 记录站点id和站点名称
        # if row[2] == "进站":
        #     row[0] = line_name + str(int(row[0]))
        #     station_id = row[0]
        #     station_name = row[1]
        #     row.append(line_name)
        #     i += 1
        #     print(row)
        #     continue
        #
        # if (row[2] == "出站") or (row[2] == "进出站"):
        #     row[0] = station_id
        #     row[1] = station_name
        #     row.append(line_name)
        #     row.append(date)
        #     data.append(row)
        #     i += 1
        #     print(row)
        #     continue
        i += 1
    # 插入数据
    with UsingMysql(log_time=True) as um:
        for line in data:
            # 把float类型转换成int类型
            for i in range(2, 16):
                line[i] = int(line[i])
        sql = "INSERT INTO line_passenger_volume_day (lineName, inPay, inNoPay, inTicket, " \
              "inCard, inTotal, turnPay, turnNoPay, turnTicket, turnCard,turnTotal, passengerPay, passengerNoPay, " \
              "passengerTicket, passengerCard, passengerTotal, countTime) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, " \
              "%s, %s, %s, %s) "
        um.cursor.executemany(sql, data)


# 只打印行列数
# def read_num(path):
#     # 打开文件
#     data = xlrd.open_workbook(path)
#     table = data.sheets()[0]
#     nrows = table.nrows
#     ncols = table.ncols
#     # print(nrows, ncols)
#     if nrows != 985 or ncols != 16:
#         print("行列数不对")
#         print(path)
#         print(nrows, ncols)


if __name__ == '__main__':

    begin = datetime.date(2021, 1, 1)
    end = datetime.date(2021, 1, 1)
    for d in range((end - begin).days + 1):
        filedate = begin + datetime.timedelta(d)
        year = str(filedate.year)
        month = str(filedate.month)
        day = str(filedate.day)
        ymd = str(filedate).replace("-", "")
        ym = ymd[0:6]
        # path = 'D:/报表/' + year + '/市地铁运营公司' + year + '-' + month + '-' + day + '/路网客运量日统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        path = 'D:/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/路网客运量日统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path = 'D:/报表/2016/北京地铁201601/北京地铁20160101/分票种进出站量日统计表(2016-1-1)-北京地铁.xls'
        # print(path)
        read_excel(path, str(filedate))
        # read_num(path)
