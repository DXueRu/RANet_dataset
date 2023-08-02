import datetime
import pandas as pd
import xlrd

from 数据代码.pymysql_lib import UsingMysql




def df_excel(path, date):
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    print(nrows, ncols)
    i = 5
    volume_list = []
    transfer_left = []
    transfer_right = []
    section_list = []
    while i < nrows:
        row = table.row_values(i)
        # print(row)
        if row[1] == '客运量(人次)':
            indicator = row[1]
            company = row[2]
            line_volume = row[12]
            row.pop(0)
            volume_list.append(row)
            while table.row_values(i + 1)[1] != '换乘站换乘量(人次)':
                i += 1
                row = table.row_values(i)
                row[1] = indicator
                row[2] = company
                row[12] = line_volume
                row.pop(0)
                volume_list.append(row)
            for line in volume_list:
                line.append(date)
                line[11] = int(line[11])
                line[12] = int(line[12])
                line[14] = int(line[14])
                for j in range(3, 10):
                    line[j] = int(line[j])

            print('客运量',volume_list)
            # todo 插入数据库

            continue
        if row[1] == '换乘站换乘量(人次)':
            indicator = row[1]
            while table.row_values(i + 1)[1] != '断面客流量(人次)':
                i += 1
                row = table.row_values(i)
                row[1] = indicator
                row.pop(0)
                row = [x for x in row if x != '']
                # print(row)
                left = row[0:7]
                transfer_left.append(left)
                if len(row) > 8:
                    right = [row[0]]
                    right += row[7:]
                    transfer_right.append(right)
            transfer_list = transfer_left + transfer_right
            # for line in transfer_list:
            #     line[1] = int(line[1])
            #     line[3] = int(line[3])
            #     line[4] = int(line[4])
            #     line.append(date)
                # print(line)
            print('换乘', transfer_list)
            # todo 插入数据库

            continue

        if row[1] == '断面客流量(人次)':
            indicator = row[1]
            company = '北京地铁'
            feature = []
            # while table.row_values(i + 1)[4] != '':
            while 1:
                row = table.row_values(i)
                if row[4] == '上行':
                    line_name = row[3]
                    section_list.append(row[3:9])
                    feature.append(row[13:])
                    i += 1
                    down = table.row_values(i)
                    down[3] = line_name
                    section_list.append(down[3:9])
                    continue
                if i == nrows - 1:
                    break
                else:
                    i += 1
            for line in section_list:
                line[2] = int(line[2])
                line.pop(4)
                line.append(date)
            for line in feature:
                line[2] = int(line[2])
                line.append(date)
            print('断面客流量', section_list)
            # todo 插入数据库
            print('出行特征', feature)
            # todo 插入数据库

            with UsingMysql(log_time=True) as um:
                # for line in feature:
                #     # 把float类型转换成int类型
                #     for i in range(3, 16):
                #         line[i] = int(float((line[i])))
                sql = "INSERT INTO Passenger_travel_characteristics_2016_2021 ( line , AverageRunningDistance_km,AverageRideStationsNumber, " \
                      "AverageTravelTime_min,countTime) VALUES ( %s, %s, %s, %s, %s) "
                um.cursor.executemany(sql, feature)

            continue

        i += 1


if __name__ == '__main__':
    begin = datetime.date(2021, 8, 19)
    end = datetime.date(2021, 11, 30)
    for d in range((end - begin).days + 1):
        filedate = begin + datetime.timedelta(d)
        year = str(filedate.year)
        month = str(filedate.month)
        day = str(filedate.day)
        ymd = str(filedate).replace("-", "")
        ym = ymd[0:6]
        # path = '/Volumes/wangyan/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ym + '\北京地铁公司'+ ymd + '\客流信息汇总表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        path = '/Volumes/wangyan/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/客流信息汇总表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path ='/Volumes/wangyan/报表/'+year+'/市地铁运营公司'+year+'-'+month+'-'+day+'/客流信息汇总表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'

        df_excel(path, str(filedate))
