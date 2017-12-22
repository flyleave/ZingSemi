import cx_Oracle
import pandas as pd


def connect2DB():
    username = "frdata"
    userpwd = "frdata2017"
    host = "10.10.17.66"
    port = 1521
    dbname = "rptdb"
    dsn = cx_Oracle.makedsn(host, port, dbname)
    connection = cx_Oracle.connect(username, userpwd, dsn)
    return connection

sql = "SELECT SUBSTR(PLAN_NAME, 5, LENGTH(PLAN_NAME)) AS PROC_EQP, A.RES_ID AS MEAS_EQP,A.SPEC_ID, B.PLAN_NAME, A.EXT_AVERAGE, A.MEAN_USL, A.MEAN_LSL, A.EXT_SIGMA, A.SIGMA_USL,A.SIGMA_LSL,A.RESULT " \
      "FROM TQS_SUMMARY_DATA@MESARCDB A, TQS_EDCPLAN_DEFINE@MESARCDB B " \
      "WHERE A.TRAN_TIME  BETWEEN to_date('2017-12-12 090000','YYYY-MM-DD hh24miss') and to_date('2017-12-19 090000','YYYY-MM-DD hh24miss') " \
      "AND A.FACTORY IN ('WE1', 'GR1') " \
      "AND A.PLAN_SEQ IN (SELECT DISTINCT PLAN_SEQ FROM TQS_EDCPLAN_DEFINE@MESARCDB WHERE PLAN_MODE='NONEMES') " \
      "AND A.PLAN_SEQ = B.PLAN_SEQ ORDER BY TRAN_TIME DESC"
connection = connect2DB()
or_df = pd.read_sql(sql, connection)

def cpk_cal(proc_eqp, spec_id, chart_type, c):
    cpk, cpu, cpl = 'Nan', 'Nan', 'Nan'
    # writer = pd.ExcelWriter('%s.xlsx' % ("lalalalalalalalala"))
    selected_df = or_df[(or_df['SPEC_ID'] == spec_id) & (or_df['PROC_EQP'] == proc_eqp)]
    # selected_df.to_excel(writer)
    if chart_type == "Xbar":
        ext_std = selected_df['EXT_AVERAGE'].std()
        ext_mean = selected_df['EXT_AVERAGE'].mean()
        print(ext_mean, ext_std)
        if (ext_std != 0) and (str(ext_std) != 'nan') and (str(ext_mean) != 'nan'):
            usl = selected_df['MEAN_USL']
            lsl = selected_df['MEAN_LSL']
            if (len(usl) != 0) & (len(lsl) != 0):
                USL = usl.iloc[0]  # first unnull value
                LSL = lsl.iloc[0]
                cpu = (USL - ext_mean) / (3 * ext_std)
                cpl = (ext_mean - LSL) / (3 * ext_std)
                cpk = min([cpu, cpl])
                print('USL LSL %s%s' % (USL, LSL))
            elif (len(usl) != 0) & (len(lsl) == 0):
                USL = usl.iloc[0]
                cpu = (USL - ext_mean) / (3 * ext_std)
                cpk = cpu
                print('222222')
            elif (len(usl) == 0) & (len(lsl) != 0):
                LSL = lsl.iloc[0]
                cpl = (ext_mean - LSL) / (3 * ext_std)
                cpk = cpl
                print('333333')
            else:
                print('都没有')

    elif chart_type == "Sigma":
        ext_std = selected_df['EXT_SIGMA'].std()
        ext_mean = selected_df['EXT_SIGMA'].mean()
        if (ext_std != 0) and (str(ext_std) != 'nan') and (str(ext_mean) != 'nan'):
            usl = selected_df['SIGMA_USL']
            lsl = selected_df['SIGMA_LSL']
            if (len(usl) != 0) & (len(lsl) != 0):
                USL = usl.iloc[0]
                LSL = lsl.iloc[0]
                cpu = (USL - ext_mean) / (3 * ext_std)
                cpl = (ext_mean - LSL) / (3 * ext_std)
                cpk = min([cpu, cpl])
            elif (len(usl) != 0) & (len(lsl) == 0):
                USL = usl.iloc[0]
                cpu = (USL - ext_mean) / (3 * ext_std)
                cpk = cpu
            elif (len(usl) == 0) & (len(lsl) != 0):
                LSL = lsl.iloc[0]
                cpl = (ext_mean - LSL) / (3 * ext_std)
                cpk = cpl
    else:
        raise Exception("Chart type '%s' does not exist" % chart_type)
    print(cpu, cpk, cpl)
    cp_dict = {'CPU': cpu, 'CPK': cpk, 'CPL': cpl}
    if c in cp_dict:
        return cp_dict[c]
    else:
        return "Not in CPU/CPK/CPL"

a= cpk_cal('AMFPI01', 'FPI-LLS065', 'Xbar', 'CPU')

def main():
    cpk_df = pd.read_excel('CPK/CPKOOC NOV.xlsx', sheetname='List')
    cpk_df['CPK值'] = cpk_df.apply(lambda _: cpk_cal(_['Collection Set'], _['MES Character ID'], 'Xbar', _['类别']), axis=1)

    writer = pd.ExcelWriter('CPK/result.xlsx')
    cpk_df.to_excel(writer, index=False)

if __name__ == '__main__':
    main()
# df = pd.DataFrame({'A':[1,2,3], 'b':[4,5,2]})
# df['a'] = df.apply(lambda _: _['A']+_['b'], axis=1)