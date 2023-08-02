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
    i = 1
    while i < nrows:
        row = table.row_values(i)


        if "分时最大断面客流量统计表" in str(row[0]):
            line_name = row[0][0:row[0].index("分时最大断面客流量统计表")]
            i += 4
            insert_data = []
            while table.row_values(i)[0] != '':
                row = table.row_values(i)
                row = [x for x in row if x != '']
                row[2] = int(row[2])
                row[5] = int(row[5])
                if row[3]=="——":
                    row[3]=0
                if row[6]=="——":
                    row[6]=0
                row.insert(0,line_name)
                row.append(date)
                # print(row)
                insert_data.append(row)
                i += 1
                if i == nrows:
                    break
            # print(insert_data)
            # todo 插入数据库
            with UsingMysql(log_time=True) as um:
                for line in insert_data:
                    print(line)
                # sql = "INSERT INTO line_section_maxvolume_hour_2021(lineName,timePeriod,section_up" \
                #     ",volume_up,ratio_up,section_down,volume_down,ratio_down,countTime) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                # um.cursor.executemany(sql,insert_data)
            continue
        i += 1


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
        # path = 'D:/报表/' + year + '/市地铁运营公司' + year + '-' + month + '-' + day + '/线路分时断面客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        path = 'D:/数据集/北京地铁数据/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/线路分时断面客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        df_excel(path, str(filedate))
