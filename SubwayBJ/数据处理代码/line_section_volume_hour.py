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

        if "分时断面客流量统计表" in str(row[0]):
            line_name = row[0][0:row[0].index("分时断面客流量统计表")]
            # 方向
            if "上行" in str(row[0]):
                direction = "上行"
            else:
                direction = "下行"
            i += 2
            row = table.row_values(i)
            if row[0] == "时间段":
                dfdata = []
                row = [x for x in row if x != '']
                dfrow = row
                dfdata.append(dfrow)
                i += 1
                while table.row_values(i)[0] != '':
                    row = table.row_values(i)
                    # 去除'' 并且 数字转换成int
                    row = [x for x in row if x != '']
                    for j in range(1, len(row)):
                        row[j] = int(row[j])

                    dfdata.append(row)
                    i += 1
                df = pd.DataFrame(dfdata, columns=dfrow)

                dfT = df.T
                # df 转为list
                ilist = dfT.values.tolist()
                del (ilist[0])
                for l in ilist:
                    l.append(line_name)
                    l.append(direction)
                    l.append(date)

                with UsingMysql(log_time=True) as um:
                    for line in ilist:
                        print(line)
                    sql = "INSERT INTO line_section_volume_hour_2021 (inter, " \
                          "T_2,T_2h,T_3,T_3h,T_4,T_4h,T_5,T_5h,T_6,T_6h," \
                          "T_7,T_7h,T_8,T_8h,T_9,T_9h,T_10,T_10h,T_11,T_11h," \
                          "T_12,T_12h,T_13,T_13h,T_14,T_14h,T_15,T_15h,T_16,T_16h," \
                          "T_17,T_17h,T_18,T_18h,T_19,T_19h,T_20,T_20h,T_21,T_21h," \
                          "T_22,T_22h,T_23,T_23h,NT_0,NT_0h,NT_1,NT_1h,lineName," \
                          "direction, countTime) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s," \
                          "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
                    um.cursor.executemany(sql, ilist)
                continue
        #         # break

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
        # path = 'D:/报表/' + year + '/市地铁运营公司' + year + '-' + month + '-' + day + '/线路分时断面客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        path = 'D:/报表/' + year + '/北京地铁公司' + ym + '/北京地铁公司' + ymd + '/线路分时断面客流量统计表(' + year + '-' + month + '-' + day + ')-北京地铁.xls'
        df_excel(path, str(filedate))
