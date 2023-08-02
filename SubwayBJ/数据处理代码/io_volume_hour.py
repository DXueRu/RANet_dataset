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
    i = 1
    # 读取数据
    while i < nrows:
        row = table.row_values(i)
        # 判段哪条线
        if "30min全票种进出站量统计表(上下行)" in str(row[0]):
            line_name = row[0][0:row[0].index("30min全票种进出站量统计表(上下行)")]
            i += 3
            continue
        # 去掉日合计和月累计
        if row[1] == "日合计":
            i += 3
            continue
        # 记录站点id和站点名称
        if row[2] == "进站":
            row[0] = line_name + str(int(row[0]))
            station_id = row[0]
            station_name = row[1]
            row.append(line_name)
            row.append(date)
            data.append(row)
            i += 1
            # print(row)
            continue

        if (row[2] == "出站") or (row[2] == "进出站"):
            row[0] = station_id
            row[1] = station_name
            row.append(line_name)
            row.append(date)
            data.append(row)
            i += 1
            # print(row)
            continue
        i += 1
    # 插入数据
    with UsingMysql(log_time=True) as um:
        for line in data:
            # 把float类型转换成int类型
            for i in range(3, 52):
                line[i] = int(line[i])
            print(line)
        sql = "INSERT INTO io_volume_hour_2021 (stationId, stationName, type, " \
              "T_2,T_2h,T_3,T_3h,T_4,T_4h,T_5,T_5h,T_6,T_6h," \
              "T_7,T_7h,T_8,T_8h,T_9,T_9h,T_10,T_10h,T_11,T_11h," \
              "T_12,T_12h,T_13,T_13h,T_14,T_14h,T_15,T_15h,T_16,T_16h," \
              "T_17,T_17h,T_18,T_18h,T_19,T_19h,T_20,T_20h,T_21,T_21h," \
              "T_22,T_22h,T_23,T_23h,NT_0,NT_0h,NT_1,NT_1h,total,"\
              "lineName, countTime) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s," \
              "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        um.cursor.executemany(sql, data)


# 只打印行列数
# def read_num(path):
#     # 打开文件
#     data = xlrd.open_workbook(path)
#     table = data.sheets()[0]
#     nrows = table.nrows
#     ncols = table.ncols
#     # print(nrows, ncols)
#     if nrows != 924 or ncols != 52:
#         print("行列数不对")
#         print(path)
#         print(nrows, ncols)


if __name__ == '__main__':

    begin = datetime.date(2021, 11, 4)
    end = datetime.date(2021, 11, 30)
    for d in range((end - begin).days + 1):
        filedate = begin + datetime.timedelta(d)
        year = str(filedate.year)
        month = str(filedate.month)
        day = str(filedate.day)
        ymd = str(filedate).replace("-", "")
        ym = ymd[0:6]
        path = 'D:/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/线路分时进出站量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path = 'D:/报表/' + year + '/市地铁运营公司' + year + '-' + month + '-' + day + '/线路分时进出站量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path = 'D:/报表/2016/北京地铁201601/北京地铁20160101/分票种进出站量日统计表(2016-1-1)-北京地铁.xls'
        # print(path)
        read_excel(path, str(filedate))
        # read_num(path)
