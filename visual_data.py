import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import pandas as pd
import datetime
import time
import cx_Oracle

username = "rptuser"
userpwd = "rptuser"
host = "10.10.17.25"
port = 1521
dbname = "mesdev"
dsn = cx_Oracle.makedsn(host, port, dbname)
connection = cx_Oracle.connect(username, userpwd, dsn)
sql = "select * from tqs_summary_data "
or_df = pd.read_sql_query(sql, connection)
print(or_df.columns)
df = or_df[['EXT_AVERAGE', 'UPDATE_TIME']]
connection.close()

#Preprocess df
# df = pd.read_excel('7050-DCOLLS037.xlsx', sheet_name=0)
df['Group'] = df['EXT_AVERAGE'].map(lambda x: '0<=X<=10'
                                    if (x >= 0 and x <= 10) else('10<X<=15' if (x <= 15 and x > 10) else 'X>=15'))
# df['UPDATE_TIME'] = df['UPDATE_TIME'].apply(lambda _: datetime.datetime.strptime(_.split(' ')[0], '%d-%b-%y'))



df = df.sort_values(by='UPDATE_TIME', ascending=True)

#select this Month
selected_df = df[df['UPDATE_TIME'].apply(lambda _:
                                         True if _.month == datetime.datetime.now().month
                                                 and _.year == datetime.datetime.now().year else False)]


#Group
grouped = selected_df.groupby([df['UPDATE_TIME'].map(lambda x: x.isocalendar()[1]), 'Group'], sort=False).agg({'EXT_AVERAGE': 'count'})

A = grouped.reset_index().pivot(index='UPDATE_TIME', columns='Group', values='EXT_AVERAGE').fillna(0)
B = A.apply(lambda _: _ * 100 / sum(_), axis=1)
ax = B.plot(kind='bar')
ax.set_xlabel("week")
ax.set_ylabel("count")
yticks = mtick.FormatStrFormatter('%.2f%%')
ax.yaxis.set_major_formatter(yticks)
ax.legend(title='')
ax.set_title('Yield of DCOLLS37 level')
for tick in ax.get_xticklabels():
    tick.set_rotation(0)
for p in ax.patches:
    ax.annotate('%.2f%%' % p.get_height(), (p.get_x() * 1.005, p.get_height() * 1.005))