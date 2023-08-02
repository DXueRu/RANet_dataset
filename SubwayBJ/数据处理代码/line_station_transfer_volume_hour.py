import datetime
import pandas as pd
import xlrd

from BJmetro.pymysql_lib import UsingMysql

def df_excel(path, date):
    #工作簿 -> 工作表 -> 单元格
    # 打开工作簿
    workbook = xlrd.open_workbook(path)
    # 获取工作表
    table = workbook.sheets()[0]
    # table = workbook.sheet_by_index(0)
    # table = workbook.sheet_by_name('Sheet1')
    nrows = table.nrows  # 工作表中的数据总行数
    ncols = table.ncols #工作表中的数据总列数
    print(nrows, ncols)

    # 初始化参数，一行一行遍历，整行的读取工作表中每一行的数据
    i = 1
    while i < nrows:
        row = table.row_values(i) #某一行的全部数据，type(row)=list
        if '分30min客流量统计表' in row[0]: #判断是否为表头
            insert_data = [] #建立列表，用于存放要导入数据库的数据
            station_name = row[0][0:row[0].index('分30min客流量统计表')] #车站名为“分”字之前的几个字符
            i += 3 #向下三行(即数据第一行)
            while table.row_values(i)[1] != '合计': #当该行第二个格的值不是“合计”时。
                row = table.row_values(i) #获取该行的全部数据
                row[0] = station_name #第一个元素为车站名
                if row[1] == '':  # 补全第二列空格为换乘量
                    row[1] = '换乘量'
                if row[1]=='换乘量' and row[2] ==''and row[3]=='':
                    i+=1
                    continue
                row.append(date) #最后一列加上日期
                insert_data.append(row) #将该行的数据列表添加进去
                i += 1
            if table.row_values(i)[1] == '合计':
                row = table.row_values(i)
                row[0] = station_name
                row.append(date)
                insert_data.append(row)
                i += 1

                # 将浮点数转为整型(第一个数字前面有一个空值，删除该空值，然后将数字转为int)
                for line in insert_data:
                    line.pop(3)
                    for j in range(3, 52):
                        line[j] = int(line[j])
                # print(insert_data)

                # todo 插入数据库
                # station_name type transferDirection
                with UsingMysql(log_time=False) as um:
                    sql = "INSERT INTO line_station_transfer_volume_hour_2021 (stationName, type1,type2, T_2, " \
                          "T_2h, T_3, T_3h, T_4, T_4h, T_5, T_5h, T_6,T_6h,T_7,T_7h,T_8,T_8h,T_9, T_9h,T_10,T_10h,T_11,T_11h,T_12,"\
                    "T_12h,T_13,T_13h,T_14,T_14h,T_15,T_15h,T_16,T_16h,T_17,T_17h,T_18,T_18h,T_19,T_19h,T_20,"\
                    "T_20h,T_21,T_21h,T_22,T_22h,T_23,T_23h,NT_0,NT_0h,NT_1,NT_1h,total,countTime) VALUES ( %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, " \
                          "%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s,%s, %s, %s, %s, %s, %s, %s, %s, %s, "\
                          "%s,%s, %s, %s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s,%s,%s) "
                    um.cursor.executemany(sql, insert_data) #批量更新

                continue #终止执行本次循环中剩下的代码，直接从下一次循环继续执行

        i += 1 #进入下一个车站的表


if __name__ == '__main__':
    # begin = datetime.date(2016, 1, 2)
    # end = datetime.date(2016, 12, 31)
    begin =datetime.date(2021,3,2)
    end = datetime.date(2021,12,31)
    for d in range((end - begin).days + 1):
        filedate = begin + datetime.timedelta(d) #timedelta表示两个date对象之间的时间间隔，精确到秒
        year = str(filedate.year)
        month = str(filedate.month)
        day = str(filedate.day)
        ymd = str(filedate).replace("-", "") #年月日
        ym = ymd[0:6] #年月
        # path = 'J:/报表/' + year + '/北京地铁' + ym + '/北京地铁' + ymd + '/换乘车站分时客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path = 'D:/PycharmProjects/BJmetro/报表/' + year + '/北京地铁' + ym + '/北京地铁' + ymd + '/换乘车站分时客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path = 'D:/PycharmProjects/BJmetro/报表/' + year + '/市地铁运营公司' + year + '-' + month + '-' + day  + '/换乘车站分时客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        path = 'D:/PycharmProjects/BJmetro/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/换乘车站分时客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'

        df_excel(path, str(filedate))
        print("已完成：",filedate)

