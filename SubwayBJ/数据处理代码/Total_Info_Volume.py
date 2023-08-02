import datetime
import pandas as pd
import xlrd



from BJmetro.pymysql_lib import UsingMysql


def df_excel(path, date):
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    print(nrows, ncols)
    i = 5
    volume_list = [] #客流信息
    while i < nrows:
        row = table.row_values(i)
        # print(row)
        if row[1] == '客运量(人次)':
            indicator = row[1] #指标
            company = row[2] #运营企业
            line_volume = row[12] #路网客运量
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
            with UsingMysql(log_time=False) as um:
                sql = "INSERT INTO Total_Info_Volume( indicator, company,linename, LineInOut," \
                      " LineInOtherOut,LineInTotal, OtherinLineOut, PassLine, TransferTotal," \
                      " DayVolume,RatioLastWeek, LineVolume, MonthVolume, RatioLastYM,YearVolume,"\
                    "RatioLastYear,countTime) VALUES ( %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, " \
                          "%s, %s, %s, %s)"
                um.cursor.executemany(sql, volume_list)
            continue
        i += 1


if __name__ == '__main__':
    begin = datetime.date(2019, 1, 1)
    end = datetime.date(2021, 11, 30)
    for d in range((end - begin).days + 1):
        filedate = begin + datetime.timedelta(d)
        year = str(filedate.year)
        month = str(filedate.month)
        day = str(filedate.day)
        ymd = str(filedate).replace("-", "")
        ym = ymd[0:6]
        path = 'D:/PycharmProjects/BJmetro/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/客流信息汇总表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        df_excel(path, str(filedate))
        print("已完成：",filedate)
