import cx_Oracle
import pandas as pd
import numpy as np
import time
import scipy.stats as st


time_range = '201711'
conf_df = pd.read_excel('CPK/Collection Character.xlsx')
block_list = "'702'"    #'711'

now_time = time.strftime('%Y-%m-%d', time.localtime(time.time()))

def connect2DB():
    username = "frdata"
    userpwd = "frdata2017"
    host = "10.10.17.66"
    port = 1521
    dbname = "rptdb"
    dsn = cx_Oracle.makedsn(host, port, dbname)
    connection = cx_Oracle.connect(username, userpwd, dsn)
    return connection


# def get_last_lot(block_list, time_range, start_spec):
#
#
#     spec_list = "'XXXX'"
#     for spec in conf_df['SPEC_ID']:
#         spec_list += ",'%s'" % spec
#     sql = "select spec_id, mat_id, substr(usr_cmf_07, 1, 10) as lot_id, substr(usr_cmf_07, 1, 8) as block_id, usr_cmf_07 as wafer_id, sys_time, ext_average " \
#           "from tqs_summary_data@Mesarcdb " \
#           "where substr(to_char(sys_time,'yyyymmdd hh24miss'), 1, 6) = '%s' " \
#           "and spec_id in (%s) and substr(usr_cmf_07, 1, 3) = %s and factory in ('WE1', 'GR1')" % (
#           time_range, spec_list, block_list)
#     # sql = "select spec_id, mat_id, lot_id, substr(LOT_ID, 1, 8) as block_id, usr_cmf_07 as wafer_id , sys_time, ext_average " \
#     #       "from tqs_summary_data@Mesarcdb " \
#     #       "where substr(to_char(sys_time,'yyyymmdd hh24miss'), 1, 6) = '%s' " \
#     #       "and substr(LOT_ID, 1, 8) in (%s)" % (time_range, block_list)
#     print(sql)
#
#     connection = connect2DB()
#     or_df = pd.read_sql(sql, connection)
#     connection.close()
#     print('Read from database sucessfully')
#
#     temp_dict = {}
#
#     index = conf_df[conf_df['SPEC_ID'] == start_spec].index.tolist()[0]
#     print(index)
#     specs = conf_df['SPEC_ID'].iloc[index: ]
#
#     for i, spec in enumerate(specs):
#         print("Current spec is : %s" % spec)
#         if i == 0:
#             temp_dict[spec] = or_df[or_df['SPEC_ID'] == spec]['LOT_ID'].unique()
#         else:
#             print("last spec: %s" % specs[index + i - 1])
#             temp_dict[spec] = list(set(or_df[or_df['SPEC_ID'] == spec]['LOT_ID'].unique()) & set(temp_dict[specs[index + i - 1]]))
#     print(temp_dict)
#     return temp_dict
#
#
#
#
#
# temp_dict = get_last_lot(block_list, '201711', '3110-A1')
#
# writer = pd.ExcelWriter('CPK/Result/last_lot.xlsx')
# temp_df = pd.DataFrame.from_dict(temp_dict, orient='index')
# temp_df.to_excel(writer, index=True)
#
#
#
# def get_raw_data(block_list, time_range):
#
#     writer = pd.ExcelWriter('CPK/Result/%s %s.xlsx' % (block_list, time_range))
#     spec_list = "'XXXX'"
#     for spec in conf_df['SPEC_ID']:
#         spec_list += ",'%s'" % spec
#     sql = "select spec_id, mat_id, lot_id, substr(LOT_ID, 1, 8) as block_id, usr_cmf_07 as wafer_id , sys_time, ext_average " \
#           "from tqs_summary_data@Mesarcdb " \
#           "where substr(to_char(sys_time,'yyyymmdd hh24miss'), 1, 6) = '%s' " \
#           "and spec_id in (%s) and substr(usr_cmf_07, 1, 3) = %s and factory in ('WE1', 'GR1')" % (time_range, spec_list, block_list)
#     # sql = "select spec_id, mat_id, lot_id, substr(LOT_ID, 1, 8) as block_id, usr_cmf_07 as wafer_id , sys_time, ext_average " \
#     #       "from tqs_summary_data@Mesarcdb " \
#     #       "where substr(to_char(sys_time,'yyyymmdd hh24miss'), 1, 6) = '%s' " \
#     #       "and substr(LOT_ID, 1, 8) in (%s)" % (time_range, block_list)
#     print(sql)
#
#     connection = connect2DB()
#     or_df = pd.read_sql(sql, connection)
#     print('Read from database sucessfully')
#
#     sta_df = pd.DataFrame()
#     col = 0
#     null_list = []
#     for spec, t in zip(conf_df['SPEC_ID'], conf_df['TYPE']):
#
#         new_df = or_df[or_df['SPEC_ID'] == spec]
#         sta_dict = {}
#         # print(new_df['SPEC_ID'].iloc[0])
#         qualified_num_df = pd.DataFrame()
#         if len(new_df) != 0:
#             sta_dict = {'SPEC': [spec], 'Average': [new_df['EXT_AVERAGE'].mean()], 'Standard Deviation': [new_df['EXT_AVERAGE'].std()],
#                         '0%': [float(np.percentile(new_df['EXT_AVERAGE'], 0))], '25%': [float(np.percentile(new_df['EXT_AVERAGE'], 25))],
#                         '50%': [float(np.percentile(new_df['EXT_AVERAGE'], 50))], '75%': [float(np.percentile(new_df['EXT_AVERAGE'], 75))],
#                         '100': [float(np.percentile(new_df['EXT_AVERAGE'], 100))]}
#         else:
#             null_list.append(spec)
#         sta_df = pd.DataFrame(sta_dict).T
#
#         if t == 'None':
#             qualified_num = len(new_df[new_df['EXT_AVERAGE'] == 0])
#             unqualified_num = len(new_df[new_df['EXT_AVERAGE'] != 0])
#             qualified_num_dict = {'日期': [time_range], '合格数量': [qualified_num], '不合格数量': [unqualified_num]}
#             qualified_num_df = pd.DataFrame(qualified_num_dict).T
#         # print(sta_df)
#         qualified_num_df.to_excel(writer, startrow=0, startcol=col, header=False)
#         sta_df.to_excel(writer, startrow=len(qualified_num_df), startcol=col, header=False)
#         new_df.to_excel(writer, startrow=len(qualified_num_df) + len(sta_df), startcol=col, index=False)
#         col += new_df.shape[1] + 1
#     print(null_list)
#     connection.close()
#
# get_raw_data(block_list, time_range)


spec_list = "'XXXX'"
for spec in conf_df['SPEC_ID']:
    spec_list += ",'%s'" % spec
sql = "select b.oper, b.recipe, a.spec_id, a.mat_id, substr(a.usr_cmf_07, 1, 10) as lot_id, substr(a.usr_cmf_07, 1, 8) as block_id, a.usr_cmf_07 as wafer_id, a.sys_time, a.ext_average " \
      "from tqs_summary_data@Mesarcdb a, MRCPMFODEF@MESARCDB b where a.mat_id = b.mat_id and a.process = b.flow and a.step_id = b.oper " \
      "and substr(to_char(a.sys_time,'yyyymmdd hh24miss'), 1, 6) = '%s' " \
      "and a.spec_id in (%s) and substr(a.usr_cmf_07, 1, 3) = %s and a.factory in ('WE1', 'GR1')" % (
          time_range, spec_list, block_list)
print(sql)
connection = connect2DB()
or_df = pd.read_sql(sql, connection)
connection.close()
print('Read from database sucessfully')
a = or_df[['WAFER_ID', 'RECIPE', 'OPER']]       # TODO RECIPE
distinct = a.drop_duplicates(['WAFER_ID', 'OPER'])
recipe_df = distinct.pivot(index='WAFER_ID', columns='OPER', values='RECIPE')


# recipe_df = pd.DataFrame(index=a['WAFER_ID'].unique(), columns=a['OPER'].unique())
# for oper in a['OPER'].unique():
#     for index in recipe_df.index.tolist():
#         recipe_df[oper].loc[index] = a[(a['OPER'] == oper) & (a['WAFER_ID'] == index)]['RECIPE'].iloc[0] if \
#             len(a[(a['OPER'] == oper) & (a['WAFER_ID'] == index)]['RECIPE']) != 0 else None



b = or_df.drop_duplicates(subset=['WAFER_ID', 'SPEC_ID'], keep='first').pivot_table(index=['MAT_ID', 'LOT_ID', 'BLOCK_ID', 'WAFER_ID', 'SYS_TIME'],
                      columns=['SPEC_ID'],
                      values='EXT_AVERAGE')
main_df = pd.merge(b.reset_index(), recipe_df.reset_index(), how='left', on='WAFER_ID')

statistic_df = pd.DataFrame(index=['PPK',
                                   'Distribution',
                                   'Process Range',
                                   'Percent Defective',
                                   'Sample Size',
                                   'Average',
                                   'Standard Deviation',
                                   '0.00%',
                                   '25.00%',
                                   '50.00%',
                                   '75.00%',
                                   '100.00%'], columns=b.columns.values.tolist())


def get_best_distribution(data):
    fitted_params_norm = st.norm.fit(data)
    fitted_params_cauchy = st.boxcox([1,2,3,4,5])
    fitted_params_expon = st.expon.fit(data)

    logLikN = np.sum(st.norm.logpdf(data, loc=fitted_params_norm[0],
    scale=fitted_params_norm[1]))


for col in b.columns.values.tolist():

    statistic_df[col].loc['PPK'] = ''
    statistic_df[col].loc['Distribution'] = ''
    statistic_df[col].loc['Process Range'] = ''
    statistic_df[col].loc['Percent Defective'] = ''
    statistic_df[col].loc['Sample Size'] = ''
    statistic_df[col].loc['Average'] = b[col].mean()
    statistic_df[col].loc['Standard Deviation'] = b[col].std()
    statistic_df[col].loc['0.00%'] = np.percentile(b[col].dropna(), 0)
    statistic_df[col].loc['25.00%'] = np.percentile(b[col].dropna(), 25)
    statistic_df[col].loc['50.00%'] = np.percentile(b[col].dropna(), 50)
    statistic_df[col].loc['75.00%'] = np.percentile(b[col].dropna(), 75)
    statistic_df[col].loc['100.00%'] = np.percentile(b[col].dropna(), 100)


writer = pd.ExcelWriter('CPK/%s Collection Data.xlsx' % now_time)
statistic_df.to_excel(writer, startcol=4)
main_df.to_excel(writer, startrow=len(statistic_df) + 1)

################
# or_df[(or_df['WAFER_ID'] == '702360380607') & (or_df['SPEC_ID'] == '6700-HazeMax')].drop_duplicates(subset='WAFER_ID', keep='first')

import scipy.stats as st
st.anderson([1,2,3,4,5,6,5,3,4,2,31,34,5,23,3,52,342,424234,52,234,2,342,34,1]) 