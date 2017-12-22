import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.ticker import FuncFormatter
import numpy as np
import pandas as pd
import datetime
import time
import os
import re
import math
import glob
import cx_Oracle
import matplotlib.gridspec as gridspec

from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart, MIMEBase
from email.utils import parseaddr, formataddr
import smtplib
import json

import xlrd
from xlwt import Style
from xlutils.copy import copy

now_time = time.strftime('%Y-%m-%d', time.localtime(time.time()))
spec_id_dict = {
    '6200-SFQRMax': '0,60,90',
    '6200-WarpBF': '0,10,17',
    '6700-LLS200': '0,5',
    '6700-AreaCount': '0,3',
    '6700-ClusterArea': '0,0.5',
    '6700-DIC': '0,3',
    '6700-ScratchLengthTotal': '0,1200',

    '7050-DCOLLS037': '0,10,15',
    '7050-DCOAreaCount': '0,1',
    '7050-ClusterArea': '0,0.5',
    '7050-DIC': '0,3',
    '7050-ScratchLengthTotal': '0,1200',
    '7050-DCNSlipLengthTotal': '0',

    '7500-DCOLLS037': '0,10,15',
    '7500-DCOAreaCount': '0,1',
    '7500-ClusterArea': '0,0.5',
    '7500-DIC': '0,3',
    '7500-ScratchLengthTotal': '0,1200',
}  # column_name_list = ['LLS037', 'LLS090', 'LLS065']


def connect2DB():
    username = "frdata"
    userpwd = "frdata2017"
    host = "10.10.17.66"
    port = 1521
    dbname = "rptdb"
    dsn = cx_Oracle.makedsn(host, port, dbname)
    connection = cx_Oracle.connect(username, userpwd, dsn)
    return connection

# def writeExcel(file, row, col, str, styl=Style.default_style):
#     rb = xlrd.open_workbook(file, formatting_info=True)
#     wb = copy(rb)
#     ws = wb.get_sheet(0)
#     ws.write(row, col, str, styl)
#     wb.save(file)

def test(con, sql):
    connection = con
    or_df = pd.read_sql(sql, connection)
    return or_df

def tqs_summary_data_excel(con, config_path):
    connection = con
    with open(config_path) as json_file:
        conf = json.load(json_file)
    select_block, select_time, select_mat = " and (1!=1", "", " and (1!=1"
    if conf['BLOCK'] != []:
        for key in conf['BLOCK']:
            select_block += " or substr(t.LOT_ID,1,8) = '%s'" % key
        select_block += ")"
        interval = 3 + len(conf['BLOCK'])
    else:
        select_block = ""
    if conf['TIME'] != []:
        select_time = " and t.TRAN_TIME between to_date('%s','YYYY-MM-DD hh24') and to_date('%s','YYYY-MM-DD hh24')" \
                      % (conf['TIME'][0], conf['TIME'][1])
    if conf['MAT'] != []:
        for key in conf['MAT']:
            select_mat += " or t.MAT_ID = %s" % key
        select_mat += ")"
        interval = 1 + len(conf['MAT'])
    else:
        select_mat = ""


    path = 'Result/%s/tqs_summary_data_yield_excel' % now_time          # output path
    loc_map_dict = {'7050': 1+2*interval, '7500': 1+3*interval, '6700': 1+interval, '6200': 1}         # excel cell row location
    count_dict = {'7050': 0, '7500': 0, '6700': 0, '6200': 0}           # excel cell columns location
    rawdata_path = 'Result/%s/tqs_summary_data_yield_excel/raw data' % now_time
    if not os.path.exists(path):
        os.makedirs(path)
    if not os.path.exists(rawdata_path):
        os.makedirs(rawdata_path)
    # select this week
    # sql = "select SPEC_ID,EXT_AVERAGE from tqs_summary_data@mesarcdb where LOT_ID like '70136029%' and UPDATE_TIME between to_date(sysdate-30) and to_date(sysdate)"
    # sql = "select SPEC_ID,EXT_AVERAGE from tqs_summary_data@mesarcdb where substr(LOT_ID,1,8) = '70136029'"
    sql = "select * " \
          "from (select t.TRAN_TIME,t.LOT_ID,t.USR_CMF_07,t.SPEC_ID,t.EXT_AVERAGE,row_number() OVER(PARTITION BY t.SPEC_ID,t.USR_CMF_07 ORDER BY t.UPDATE_TIME desc) as row_flag from tqs_summary_data@mesarcdb t where 1=1%s%s%s) temp " \
          "where temp.row_flag='1'" % (select_block, select_time, select_mat)
    # sql = "select * from tqs_summary_data@mesarcdb " \
    #       "where retest_step = 0 and (%s)%s" % (select_block, select_time)

    # sql = "select * from tqs_summary_data@mesarcdb where retest_step=0%s%s" % (select_block, select_time)
    print("select criteria: %s" % sql)
    or_df = pd.read_sql(sql, connection)
    print("selected %s records" % len(or_df))

    excel_path = '%s/%s Yield.xlsx' % (path, now_time)
    writer = pd.ExcelWriter(excel_path)
    t_6200 = {"BLOCK_ID": conf['BLOCK'], "TOTAL": []}
    t_6700 = {"BLOCK_ID": conf['BLOCK'], "TOTAL": []}
    t_7050 = {"BLOCK_ID": conf['BLOCK'], "TOTAL": []}
    t_7500 = {"BLOCK_ID": conf['BLOCK'], "TOTAL": []}
    for key in spec_id_dict:
        proc_name = key.split('-')[0]
        condiction = key.split('-')[1]
        total_df = or_df[or_df['SPEC_ID'] == key]
        main_df = pd.DataFrame()
        l = []

        for block in conf['BLOCK']:
            if len(total_df) == 0:
                l.append(0)
                continue
            print("SPEC is %s, BLOCK is %s" % (key, block))
            new_df = total_df[total_df['LOT_ID'].apply(lambda _: True if re.match(block, str(_)) else False)]
            print("total_df is:")
            print(total_df)
            print("new_df is:")
            print(new_df)
            xiaoyan_writer = pd.ExcelWriter('%s/%s-%s.xlsx' % (rawdata_path, key, block))
            pd.DataFrame(new_df).to_excel(xiaoyan_writer)
            l.append(len(new_df))
            if spec_id_dict[key] == "0,10,15":    # 0<=x<=10  10<x<=15  x>=15
                temp_dict = {key: ['0<=X<=10', '10<X<=15', 'X>15'],
                             'Count': [len(new_df[(new_df['EXT_AVERAGE'] <= 10) & (new_df['EXT_AVERAGE'] >= 0)]),
                                       len(new_df[(new_df['EXT_AVERAGE'] <= 15) & (new_df['EXT_AVERAGE'] > 10)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 15])]
                             }
            elif spec_id_dict[key] == "0,1,2":
                # new_df['Group'] = new_df['EXT_AVERAGE'].map(
                #     lambda x: '0' if x == 0 else('1' if x == 1 else('2' if x == 2 else '>2')))
                temp_dict = {key: ['X=0', '0<X<=2', 'X>=2'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[(new_df['EXT_AVERAGE'] <= 2) & (new_df['EXT_AVERAGE'] > 0)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 2])]
                             }
            elif spec_id_dict[key] == "0,5,10":
                temp_dict = {key: ['0<=X<=5', '5<X<=10', 'X>10'],
                             'Count': [len(new_df[(new_df['EXT_AVERAGE'] >= 0) & (new_df['EXT_AVERAGE'] <= 5)]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 5) & (new_df['EXT_AVERAGE'] <= 10)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 10])]
                             }
            elif spec_id_dict[key] == "0,1200":
                temp_dict = {key: ['X=0', '0<X<=1.2', 'X>1.2'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 0) & (new_df['EXT_AVERAGE'] <= 1200)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 1200])]
                             }
            elif spec_id_dict[key] == "0,0.5":
                # new_df['Group'] = new_df['EXT_AVERAGE'].map(
                #     lambda x: '0' if x == 0 else('0<X<=1.5' if (x <= 1.5 and x > 0) else 'X>1.5'))
                temp_dict = {key: ['X=0', '0<X<=0.5', 'X>0.5'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 0) & (new_df['EXT_AVERAGE'] <= 0.5)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 0.5])]
                             }
            elif spec_id_dict[key] == "0,5":
                # new_df['Group'] = new_df['EXT_AVERAGE'].map(
                #     lambda x: '0' if x == 0 else('0<X<=1.5' if (x <= 1.5 and x > 0) else 'X>1.5'))
                temp_dict = {key: ['X=0', '0<X<=5', 'X>5'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 0) & (new_df['EXT_AVERAGE'] <= 5)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 5])]
                             }
            elif spec_id_dict[key] == "0,3":
                temp_dict = {key: ['X=0', '0<X<=3', 'X>3'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 0) & (new_df['EXT_AVERAGE'] <= 3)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 3])]
                             }
            elif spec_id_dict[key] == "0":
                # new_df['Group'] = new_df['EXT_AVERAGE'].map(
                #     lambda x: '0' if x == 0 else ">0")
                temp_dict = {key: ['X=0', 'X>0'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 0])]
                             }
            elif spec_id_dict[key] == "0,60,90":
                # new_df['Group'] = new_df['EXT_AVERAGE'].map(
                #     lambda x: '0' if x == 0 else ">0")
                temp_dict = {key: ['0<=X<=60', '60<X<=90', 'X>90'],
                             'Count': [len(new_df[(new_df['EXT_AVERAGE'] >= 0) & (new_df['EXT_AVERAGE'] <= 60)]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 60) & (new_df['EXT_AVERAGE'] <= 90)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 90])]
                             }
            elif spec_id_dict[key] == "0,10,17":
                # new_df['Group'] = new_df['EXT_AVERAGE'].map(
                #     lambda x: '0' if x == 0 else ">0")
                temp_dict = {key: ['0<=X<=10', '10<X<=17', 'X>17'],
                             'Count': [len(new_df[(new_df['EXT_AVERAGE'] >= 0) & (new_df['EXT_AVERAGE'] <= 10000)]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 10000) & (new_df['EXT_AVERAGE'] <= 17000)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 17000])]
                             }
            elif spec_id_dict[key] == "0,1":
                temp_dict = {key: ['X=0', '0<X<=1', 'X>1'],
                             'Count': [len(new_df[new_df['EXT_AVERAGE'] == 0]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 0) & (new_df['EXT_AVERAGE'] <= 1)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 1])]
                             }

            elif spec_id_dict[key] == "0,60,90":
                temp_dict = {key: ['0<=X<=60', '60<X<=90', 'X>90'],
                             'Count': [len(new_df[(new_df['EXT_AVERAGE'] >= 0) & (new_df['EXT_AVERAGE'] <= 60)]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 60) & (new_df['EXT_AVERAGE'] <= 90)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 90])]
                             }

            elif spec_id_dict[key] == "0,10,17":
                temp_dict = {key: ['0<=X<=10', '10<X<=17', 'X>17'],
                             'Count': [len(new_df[(new_df['EXT_AVERAGE'] >= 0) & (new_df['EXT_AVERAGE'] <= 10000)]),
                                       len(new_df[(new_df['EXT_AVERAGE'] > 10000) & (new_df['EXT_AVERAGE'] <= 17000)]),
                                       len(new_df[new_df['EXT_AVERAGE'] > 17000])]
                             }

            else:
                raise Exception("No such range!!!: %s" % spec_id_dict[key])
            # print(temp_dict)
            rs_df = pd.DataFrame(temp_dict)
            # print(rs_df['Count'].sum())
            if rs_df['Count'].sum() != 0:
                rs_df['Count'] = 100 * rs_df['Count']/rs_df['Count'].sum()        # 转化比例
                rs_df['Count'] = rs_df['Count'].apply(lambda _: "%.2f%%" % _)       #转化百分数
            else:
                rs_df['Count'] = "NA"
            a = rs_df.set_index(key)
            b = a.T
            main_df = pd.concat([main_df, b])

        # print("proc_name: %s" % proc_name)
        # print(l)
        if proc_name == "6700":
            t_6700['TOTAL'] = l
        elif proc_name == "7050":
            t_7050['TOTAL'] = l
        elif proc_name == "7500":
            t_7500['TOTAL'] = l
        elif proc_name == "6200":
            t_6200['TOTAL'] = l



        # grouped = new_df.groupby(['Group'], sort=False).agg({'EXT_AVERAGE': 'count'})
        # A = grouped.reset_index()
        # A.rename(columns={'EXT_AVERAGE': 'count', 'Group': condiction}, inplace=True)
        # A['yield'] = A['count']/A['count'].sum()



        # print(A)
        # print(loc_map_dict[proc_name], (count_dict[proc_name] - 1) * 3)
        # print(A, loc_map_dict[proc_name], (count_dict[proc_name] - 1) * 3)

        t = pd.DataFrame({key: []})
        t.to_excel(writer, startrow=loc_map_dict[proc_name] - 1, startcol=2 + count_dict[proc_name], index=False)
        main_df.to_excel(writer, startrow=loc_map_dict[proc_name], startcol=2 + count_dict[proc_name], index=False)

        count_dict[proc_name] += main_df.shape[1]
# order is: 6200,6700,7050,7500
    pd.DataFrame(t_6200).to_excel(writer, startrow=1, startcol=0, index=False)
    pd.DataFrame(t_6700).to_excel(writer, startrow=1+interval, startcol=0, index=False)
    pd.DataFrame(t_7050).to_excel(writer, startrow=1+2*interval, startcol=0, index=False)
    pd.DataFrame(t_7500).to_excel(writer, startrow=1+3*interval, startcol=0, index=False)

    cnt_7500 = 0    # 7500 spec的数量
    cnt_7500_wafer = 0         # 7500 合格wafer的数量
    qualified_dict = {'count': [], 'yield': []}
    a_writer = pd.ExcelWriter('%s/%s.xlsx' % (rawdata_path, "oooooooo"))
    or_df['Flag'] = 0
    for spec in spec_id_dict:
        if spec.split('-')[0] == '7500':
            upline = float(spec_id_dict[spec].split(',')[-1])
            print("upline: %s" % upline)

            or_df['Flag'] = or_df.apply(lambda _: 1 if ((_['SPEC_ID'] == spec) & (_['EXT_AVERAGE'] <= upline)) or (_['Flag'] == 1) else 0, axis=1)
            # or_df[(or_df['SPEC_ID'] == spec) & (or_df['EXT_AVERAGE'] <= upline)]['Flag'] = 1
            cnt_7500 += 1
    print(or_df)
    or_df.to_excel(a_writer)
    print("cnt_7500: %s" % cnt_7500)
    for block in conf['BLOCK']:
        cnt_7500_wafer = 0
        new_df = or_df[or_df['LOT_ID'].apply(lambda _: True if re.match(block, str(_)) else False)]
        print(new_df)
        for wafer in new_df['USR_CMF_07'].unique():
            if len(new_df[(new_df['USR_CMF_07'] == wafer) & (new_df['Flag'] == 1)]) == cnt_7500:
                cnt_7500_wafer += 1
        qualified_dict['count'].append(cnt_7500_wafer)
        if len(new_df[new_df['SPEC_ID'] == '7500-DIC']) != 0:
            qualified_dict['yield'].append("%.2f%%" % (100 * cnt_7500_wafer / len(new_df[new_df['SPEC_ID'] == '7500-DIC'])))
        else:
            qualified_dict['yield'].append('NA')

    pd.DataFrame(qualified_dict).to_excel(writer, startrow=loc_map_dict["7500"], startcol=2 + count_dict["7500"], index=False)



# def plot_tqs_summary_data_df(con, spec_id_dict):
#     '''
#     :param con: connection with oracle
#     :param title: a dict, key is spec_id and value is criterion
#                                 which contains: 1."0,10,15", 2."0,1,2", 3."0,5,10", 4."0,1.2", 5."0,1.5", 6."0,1,2,3"
#     :return: none
#     '''
#     connection = con
#     if not os.path.exists('Result/%s/yield_bar_plot' % now_time):
#         os.makedirs('Result/%s/yield_bar_plot' % now_time)
#     # Select by title
#     for key in spec_id_dict:
#         sql = "select * from tqs_summary_data@mesarcdb where SPEC_ID=" + "'" + key + "'"
#         print("Select criterion is: %s" % sql)
#         or_df = pd.read_sql(sql, connection)
#
#         df1 = or_df[['EXT_AVERAGE', 'UPDATE_TIME']]
#         if len(df1) == 0:
#             print("0 rows is selected")
#         else:
#             # Preprocessing df
#             df = df1.copy()
#             if spec_id_dict[key] == "0,10,15":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0<=X<=10' if (x >= 0 and x <= 10) else('10<X<=15' if (x <= 15 and x > 10) else 'X>=15'))
#             elif spec_id_dict[key] == "0,1,2":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0' if x == 0 else('1' if x == 1 else('2' if x == 2 else '>2')))
#             elif spec_id_dict[key] == "0,5,10":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0<=X<=5' if (x >= 0 and x <= 10) else('5<X<=10' if (x <= 10 and x > 5) else 'X>10'))
#             elif spec_id_dict[key] == "0,1.2":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0' if x == 0 else('0<X<=1.2' if (x <= 1.2 and x > 0) else 'X>1.2'))
#             elif spec_id_dict[key] == "0,1.5":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0' if x == 0 else('0<X<=1.5' if (x <= 1.5 and x > 0) else 'X>1.5'))
#             elif spec_id_dict[key] == "0,1,2,3":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0' if x == 0 else('1' if x == 1 else('2' if x == 2 else ('3' if x == 3 else '>3'))))
#             elif spec_id_dict[key] == "0,0.5":
#                 df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0' if x == 0 else('0<X<=0.5' if (x <= 0.5 and x > 0) else 'X>0.5'))
#
#             else:
#                 raise Exception("No such criterion: %s" % spec_id_dict[key])
#             df = df.sort_values(by='UPDATE_TIME', ascending=True)
#
#             # Select this Month
#             selected_df = df[df['UPDATE_TIME'].apply(lambda _:
#                                                      True if _.month == datetime.datetime.now().month
#                                                              and _.year == datetime.datetime.now().year else False)]
#             if len(selected_df) == 0:
#                 print("No data this month for %s" % key)
#             else:
#                 # Group
#                 grouped = selected_df.groupby([df['UPDATE_TIME'].map(lambda x: x.isocalendar()[1]), 'Group'], sort=False).agg({'EXT_AVERAGE': 'count'})
#
#                 A = grouped.reset_index().pivot(index='UPDATE_TIME', columns='Group', values='EXT_AVERAGE').fillna(0)
#                 B = A.apply(lambda _: _ * 100 / sum(_), axis=1)
#                 ax = B.plot(kind='bar')
#                 ax.set_xlabel("week")
#                 ax.set_ylabel("count")
#                 yticks = mtick.FormatStrFormatter('%.2f%%')
#                 ax.yaxis.set_major_formatter(yticks)
#                 ax.legend(title='')
#                 ax.set_title(key)
#                 for tick in ax.get_xticklabels():
#                     tick.set_rotation(0)
#                 for p in ax.patches:
#                     ax.annotate('%.2f%%' % p.get_height(), (p.get_x() * 1.005, p.get_height() * 1.005))
#
#                 # Save Image
#                 fig = ax.get_figure()
#                 fig.savefig('Result/%s/yield_bar_plot/%s.png' % (now_time, key))

def plot_bar(con, config_path, map_path):
    path = "Result/%s/TM_BAR" % now_time
    select_time = ""
    select_mat = ""
    if not os.path.exists(path):
        os.makedirs(path)
    connection = con
    with open(config_path) as json_file:
        conf = json.load(json_file)
    if conf['MAT'] != []:
        select_mat = " and ("
        for key in conf['MAT']:
            select_mat += " t.MAT_ID='%s' or" % key
        select_mat += " 1!=1)"
    if conf['TIME'] != []:
        select_time = " and t.TRAN_TIME between to_date('%s','YYYY-MM-DD hh24') and to_date('%s','YYYY-MM-DD hh24')" \
                      % (conf['TIME'][0], conf['TIME'][1])

    sql = "select temp.SPEC_ID,temp.EXT_AVERAGE " \
        "from (select t.TRAN_TIME,t.LOT_ID,t.SPEC_ID,t.EXT_AVERAGE,row_number() OVER(PARTITION BY t.SPEC_ID,t.USR_CMF_07 ORDER BY t.UPDATE_TIME desc) as row_flag from tqs_summary_data@mesarcdb t where length(t.USR_CMF_07)=12%s%s) temp " \
        "where temp.row_flag='1'" % (select_time, select_mat)

    # sql = "select SPEC_ID, EXT_AVERAGE from tqs_summary_data@mesarcdb where length(t.USR_CMF_07)=12%s%s and retest_step=0" % (select_time, select_mat)
    print("sql = %s" % sql)
    or_df = pd.read_sql(sql, connection)
    #####################
    writer = pd.ExcelWriter('%s/or_df.xlsx' % path)
    or_df.to_excel(writer, startrow=0, startcol=0)
    ##################
    map_df = pd.read_excel(map_path)
    TM37_dict = {}
    TM65_dict = {}
    TM90_dict = {}
    for i, j, bottom, top in zip(map_df['EQP'], map_df['SPEC_ID'], map_df['TM37_BOTTOM'], map_df['TM37_TOP']):
        spec_id = "%s-%s" % (i, j)
        if (not pd.isnull(bottom) or not pd.isnull(top)) and len(or_df[(or_df['SPEC_ID'] == spec_id)]) == 0:
            TM37_dict[spec_id] = 0
        else:
            if not pd.isnull(bottom) and not pd.isnull(top):
                TM37_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] <= top)
                                               & (or_df['EXT_AVERAGE'] >= bottom)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            elif pd.isnull(bottom) and not pd.isnull(top):
                TM37_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] <= top)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            elif not pd.isnull(bottom) and pd.isnull(top):
                TM37_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] >= bottom)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            else:
                continue
    for i, j, bottom, top in zip(map_df['EQP'], map_df['SPEC_ID'], map_df['TM65_BOTTOM'], map_df['TM65_TOP']):
        spec_id = "%s-%s" % (i, j)
        if (not pd.isnull(bottom) or not pd.isnull(top)) and len(or_df[(or_df['SPEC_ID'] == spec_id)]) == 0:
            TM65_dict[spec_id] = 0
        else:
            if not pd.isnull(bottom) and not pd.isnull(top):
                TM65_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] <= top)
                                               & (or_df['EXT_AVERAGE'] >= bottom)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            elif pd.isnull(bottom) and not pd.isnull(top):
                TM65_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] <= top)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            elif not pd.isnull(bottom) and pd.isnull(top):
                TM65_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] >= bottom)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            else:
                continue
    for i, j, bottom, top in zip(map_df['EQP'], map_df['SPEC_ID'], map_df['TM90_BOTTOM'], map_df['TM90_TOP']):
        spec_id = "%s-%s" % (i, j)
        if (not pd.isnull(bottom) or not pd.isnull(top)) and len(or_df[(or_df['SPEC_ID'] == spec_id)]) == 0:
            TM90_dict[spec_id] = 0
        else:
            if not pd.isnull(bottom) and not pd.isnull(top):
                TM90_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] <= top)
                                               & (or_df['EXT_AVERAGE'] >= bottom)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            elif pd.isnull(bottom) and not pd.isnull(top):
                TM90_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] <= top)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            elif not pd.isnull(bottom) and pd.isnull(top):
                TM90_dict[spec_id] = len(or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['EXT_AVERAGE'] >= bottom)]) / len(or_df[or_df['SPEC_ID'] == spec_id])
            else:
                continue
    print(TM37_dict, TM65_dict, TM90_dict) # TODO sort the 3 dictionary
    labels_37 = [x for x in TM37_dict.keys()]
    values_37 = [x * 100 for x in TM37_dict.values()]
    labels_65 = [x for x in TM65_dict.keys()]
    values_65 = [x * 100 for x in TM65_dict.values()]
    labels_90 = [x for x in TM90_dict.keys()]
    values_90 = [x * 100 for x in TM90_dict.values()]
    print("TM37:")
    print(labels_37)
    print(values_37)
    # plot figure
    yticks = mtick.FormatStrFormatter('%.0f%%')


    fig1 = plt.figure(1)
    ax1 = fig1.add_subplot(1, 1, 1)
    ax1.bar(labels_37, values_37, 0.4, color="green")
    ax1.yaxis.set_major_formatter(yticks)
    for a, b in zip(labels_37, values_37):
        ax1.text(a, b+1, '%.0f%%' % b, ha='center', va='center')
    ax1.set_title("TM37")
    ax1.set_ylabel("Yield")
    fig1.set_size_inches(16, 12)
    plt.xticks(rotation=45)
    fig1.savefig("%s/TM37.jpg" % path)

    fig2 = plt.figure(2)
    ax2 = fig2.add_subplot(1, 1, 1)
    ax2.bar(labels_65, values_65, 0.4, color="green")
    ax2.yaxis.set_major_formatter(yticks)
    for a, b in zip(labels_65, values_65):
        ax2.text(a, b+1, '%.0f%%' % b, ha='center', va='center')
    ax2.set_title("TM65")
    ax2.set_ylabel("Yield")
    fig2.set_size_inches(16, 12)
    plt.xticks(rotation=45)
    fig2.savefig("%s/TM65.jpg" % path)

    fig3 = plt.figure(3)
    ax3 = fig3.add_subplot(1, 1, 1)
    ax3.bar(labels_90, values_90, 0.4, color="green")
    ax3.yaxis.set_major_formatter(yticks)
    for a, b in zip(labels_90, values_90):
        ax3.text(a, b+1, '%.0f%%' % b, ha='center', va='center')
    ax3.set_title("TM90")
    ax3.set_ylabel("Yield")
    fig3.set_size_inches(16, 12)
    plt.xticks(rotation=45)
    fig3.savefig("%s/TM90.jpg" % path)

def plot_box(con, config_path):
    '''
    :param con: connection with oracle
    :param colname_list: list of columns to plot
    :return: none
    '''
    with open(config_path) as json_file:
        conf = json.load(json_file)
    select_block, select_time, select_EQP, select_MAT_ID = "", "", "", ""
    select_cols = ",%s" % conf['columns']     # conf['columns'] = "a,b,c,d"

    if conf['MAT'] != []:  # conf['MAT'] = []
        select_MAT_ID = " and (1!=1"
        for key in conf['MAT']:
            select_MAT_ID += " or MAT_ID = '%s'" % key
        select_MAT_ID += ")"

    if conf['EQP'] != "":
        select_EQP = " and (1!=1"
        for key in conf['EQP']:
            select_block += " or PROC_EQP = '%s'" % key
        select_block += ")"

    if conf['TIME'] != []:
        select_time = " and SPX_Time between to_date('%s','YYYY-MM-DD hh24') and to_date('%s','YYYY-MM-DD hh24')" \
                      % (conf['TIME'][0], conf['TIME'][1])

    if conf['BLOCK'] != []:
        select_block = " and (1!=1"
        for key in conf['BLOCK']:
            select_block += " or BLOCK_ID = '%s'" % key
        select_block += ")"

    sql = "select to_char(to_date(SPX_TIME,'yyyymmddhh24miss'),'iw') as Week, PROC_EQP %s " \
              "from VSP3FNPYLD where 1=1%s%s%s%s" % (select_cols, select_MAT_ID, select_EQP, select_time, select_block)

    print("Select criterion is: %s" % sql)
    connection = con
    if not os.path.exists('Result/%s/box_plot' % now_time):
        os.makedirs('Result/%s/box_plot' % now_time)


    sql_df = pd.read_sql(sql, connection)
    for colname in conf['columns'].split(','):
        or_df = sql_df[['WEEK', 'PROC_EQP', colname]]
        proc_list = or_df['PROC_EQP'].unique()
        proc_list.sort()
        fig = plt.figure(colname)
        N = len(proc_list)
        cols = 2
        rows = int(math.ceil(N / cols))
        gs = gridspec.GridSpec(rows, cols)
        for n, proc_id in enumerate(proc_list):
            ax = fig.add_subplot(gs[n])
            sub_df = or_df[or_df['PROC_EQP'] == proc_list[n]]
            plot_dict = sub_df.boxplot(column=colname, by=['WEEK'], ax=ax, return_type='dict', showfliers=False)
            ax.set_title(proc_id)
            m1 = sub_df.groupby(['WEEK'])[colname].median().values
            mL1 = [str(np.round(s, 2)) for s in m1]
            for tick in range(len(m1)):
                ax.text(tick + 1, m1[tick] + 10, mL1[tick], horizontalalignment='center', color='b', weight='semibold')
            x = np.linspace(1, len(m1), len(m1))
            y = np.array(m1)
            ax.plot(x, y)
            # ax.set_ylim([-50, 1000])
            ax.set_xlabel('Wxx, 周')
            # TODO fill box color
            # for patch in plot_dict['boxes']:
            #     patch.set_facecolor('lightblue')
            fig.suptitle(colname)
            # Save figure
            fig.set_size_inches(15,10)
            fig.savefig('Result/%s/box_plot/%s.png' % (now_time, colname))

# Send email
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def send_email(from_addr, pwd, to_addr, subject, pics_path):

    msg = MIMEMultipart()
    msg['From'] = _format_addr('<%s>' % from_addr)
    msg['To'] = _format_addr('<%s>' % to_addr)
    msg['Subject'] = Header(subject, 'utf-8').encode()
    msg.attach(MIMEText('Text', 'plain', 'utf-8'))
    for pic in pics_path:
        pic_name = pic.split('/')[1]
        print("Adding %s to attachment" % pic_name)
        with open(pic, 'rb') as f:
            mime = MIMEBase('image', 'png', filename=pic_name)
            mime.add_header('Content-Disposition', 'attachment', filename=pic_name)
            mime.add_header('Content-ID', '<0>')
            mime.add_header('X-Attachment-Id', '0')
            mime.set_payload(f.read())
            encoders.encode_base64(mime)
            msg.attach(mime)

    smtp_server = 'owa.zingsemi.com'

    server = smtplib.SMTP(smtp_server, 25)
    server.set_debuglevel(1)
    server.login(from_addr, pwd)
    server.sendmail(from_addr, [to_addr], msg)
    server.quit()

# zingsemi\e00xxx
def main():
    # from_addr = "melody.he@zingsemi.com"      # sender
    # pwd = "Xx147258"            # password
    # to_addr = "melody.he@zingsemi.com"        # receiver
    # subject = "Yield"   # subject

    connection = connect2DB()
    # plot_tqs_summary_data_df(connection, spec_id_dict)
    # plot_bar(connection, config_path='bar_tm_config.json', map_path="mapping.xlsx")
    tqs_summary_data_excel(connection, config_path='tqs_excel_config.json')

    # plot_box(connection, config_path='box_config.json')
    # pngList = glob.glob('Result/%s/*.png' % now_time)
    # print("Preparing sending %d pics" % len(pngList))
    # send_email(from_addr, pwd, to_addr, subject, pngList)
    connection.close()


if __name__ == '__main__':
    main()
# connection = connect2DB()
# sql = "select temp.USR_CMF_07,temp.SPEC_ID,temp.EXT_AVERAGE from (select t.USR_CMF_07,t.TRAN_TIME,t.LOT_ID,t.SPEC_ID,t.EXT_AVERAGE,row_number() OVER(PARTITION BY t.SPEC_ID,t.USR_CMF_07 ORDER BY t.UPDATE_TIME desc) as row_flag from tqs_summary_data@mesarcdb t) temp where temp.row_flag='1' and temp.TRAN_TIME between to_date('2017-10-01','YYYY-MM-DD') and to_date('2017-11-28','YYYY-MM-DD') and substr(temp.LOT_ID,1,8) = '701390F5'"
# # sql = "select t.USR_CMF_07,t.TRAN_TIME,t.SPEC_ID,t.EXT_AVERAGE,row_number() OVER(PARTITION BY t.SPEC_ID,t.USR_CMF_07 ORDER BY t.UPDATE_TIME desc) as row_flag from tqs_summary_data@mesarcdb t where t.TRAN_TIME between to_date('2017-10-01','YYYY-MM-DD') and to_date('2017-11-28','YYYY-MM-DD') and substr(t.LOT_ID,1,8) = '701390F5'"
# a = test(connection, sql)