import datetime
import pandas as pd
import xlrd


from pymysql_lib import UsingMysql





def df_excel(path, date):
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    print(nrows, ncols)
    i = 5
    transfer_left = []
    transfer_right = []
    while i < nrows:
        row = table.row_values(i)
        if row[1] == '换乘站换乘量(人次)':
            indicator = row[1]
            while table.row_values(i+1)[1] != '断面客流量(人次)':
                i += 1
                row = table.row_values(i)
                row[1] = indicator
                row.pop(0)
                row = [x for x in row if x != '']
                left = row[0:7]
                transfer_left.append(left)
                if len(row) > 8:
                    right = [row[0]]
                    right += row[7:]
                    transfer_right.append(right)
            transfer_list = transfer_left + transfer_right
            for line in transfer_list:
                line[1] = int(line[1])
                line[3] = int(line[3])
                line[4] = int(line[4])
                line.append(date)
                # print(line)
            # todo 插入数据库
            with UsingMysql(log_time=True) as um:
                for line in transfer_list:
                    print(line)
                sql = "INSERT INTO Total_Info_Transfer_2016_2021 ( indicator, sort, TransferStation, DayTransferVolume, " \
                        "HourMaxTransferVolume, timePeriod, RatioAllDay, countTime) VALUES ( %s, %s, %s, %s, %s, %s, %s, %s) "
                um.cursor.executemany(sql, transfer_list)

            continue

        i += 1


if __name__ == '__main__':
    begin = datetime.date(2021, 1, 1)
    end = datetime.date(2021, 11, 30)
    for d in range((end - begin).days + 1):
        filedate = begin + datetime.timedelta(d)
        year = str(filedate.year)
        month = str(filedate.month)
        day = str(filedate.day)
        ymd = str(filedate).replace("-", "")
        ym = ymd[0:6]
        path = 'D:/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/客流信息汇总表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        # path = 'D:/报表/' + year + '/市地铁运营公司' + year + '-' + month + '-' + day + '/客流信息汇总表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        df_excel(path, str(filedate))
