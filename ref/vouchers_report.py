#  Copyright (c) 2021. Alexander Zarkov
import warnings
# import pymysql
import sys
import pandas as pd
from sqlalchemy import create_engine
from pandasql import sqldf

sys.path.insert(0, 'c:/xampp/htdocs/includes/')
import webdb
import send_mail

warnings.filterwarnings("ignore")
pd.set_option('display.max_columns', 20)
pd.set_option('display.width', 300)

query = """
    SELECT
        l.user_id,
        l.used_on,
        v. NAME AS voucher_name,
        count(l.voucher_id) AS 'count'
    FROM
        vouchers_log l
    INNER JOIN vouchers_codes c ON l.voucher_id = c.id
    AND l.user_id = c.user_id
    INNER JOIN vouchers v ON v.voucher_code = c.voucher_code
    WHERE
        l.generated_on >= CURDATE() - INTERVAL 3 MONTH
    GROUP BY
        l.used_on,
        v. NAME
    HAVING
        count(l.voucher_id) > 1
    ORDER BY
        l.user_id,
        count DESC
"""

engine = create_engine(f'mysql+pymysql://{webdb.user}:{webdb.pw}@{webdb.host}:3306/mtelbg_club', echo=False)
db = engine.connect()
df_v = pd.read_sql(query, con=engine)

query = """
    SELECT
        DATE_FORMAT(l.generated_on, '%%Y-%%m') AS months,
        v. NAME AS voucher_name,
        l.used_on,
        l.user_id
    FROM
        vouchers_log l
    INNER JOIN vouchers_codes c ON l.voucher_id = c.id
    AND l.user_id = c.user_id
    INNER JOIN vouchers v ON v.voucher_code = c.voucher_code
    WHERE
        l.generated_on >= CURDATE() - INTERVAL 3 MONTH
    ORDER BY
        l.used_on
"""
df_act = pd.read_sql(query, con=engine)
print(df_act.head())
db.close()

z = sqldf("select distinct t.months, t.voucher_name, t.used_on, t.user_id from df_act t inner join df_act t1 "
          "where t.months <> t1.months and t.voucher_name = t1.voucher_name "
          "and t.used_on = t1.used_on and t.user_id <> t1.user_id")
z['domain_user'] = z.user_id + '@a1.bg'
print(z.head())

# print(df_v.head())
df_v['domain_user'] = df_v.user_id + '@a1.bg'
print(df_v.head())

# users = "'" + "','".join(list(set(df_v['domain_user'].tolist()))) + "'"
# print(users)

query = f"""
    SELECT
        ed.NamesBG,
        ed.DomainUser,
        p.NameBG as Position,
        sd2.NameBG as Direction,
        sd1.NameBG as Department,
        sd.NameBG as Team,
        e.EmailAddress
    FROM
        `EmployeeData` ed
    INNER JOIN Emails e ON e.EmployeeID = ed.EmployeeID
    INNER JOIN EmployeeInStruct eis ON eis.EmployeeID = ed.EmployeeID
    INNER JOIN StructData sd ON eis.StructID = sd.StructID
    INNER JOIN StructData sd1 ON sd.ParentID = sd1.StructID
    INNER JOIN StructData sd2 ON sd1.ParentID = sd2.StructID
    INNER JOIN Positions p ON p.PositionID = eis.Position
"""

# print(query)
engine = create_engine(f'mysql+pymysql://{webdb.user}:{webdb.pw}@{webdb.host}:3306/Staff', echo=False)
db = engine.connect()
df_u = pd.read_sql(query, con=engine)
db.close()

print(df_u.head())

df_final = pd.merge(df_v, df_u, how="left", left_on="domain_user", right_on='DomainUser')
df_final[['used_on']] = df_final[['used_on']].astype('str')

z = pd.merge(z, df_u, how="left", left_on="domain_user", right_on='DomainUser')
z[['used_on']] = z[['used_on']].astype('str')

print(df_final.head())
df_final.to_csv('df_final.csv', index=False, encoding='utf-8-sig', sep=';')

d = df_final[['used_on', 'count', 'user_id', 'voucher_name']].groupby(['used_on'], sort=True). \
    agg({'count': 'sum', 'user_id': 'nunique', 'voucher_name': 'nunique'}). \
    sort_values('count', ascending=False)

print(d.head())
# d.to_csv('counts.csv', index=True, encoding='utf-8-sig', sep=';')

shops = df_final[['Team', 'count', 'user_id', 'voucher_name']].groupby(['Team'], sort=True). \
    agg({'count': 'sum', 'user_id': 'nunique', 'voucher_name': 'nunique'}). \
    sort_values('count', ascending=False)
print(shops.head())

users = df_final[['NamesBG', 'count', 'used_on', 'voucher_name']].groupby(['NamesBG'], sort=True). \
    agg({'count': 'sum', 'used_on': 'nunique', 'voucher_name': 'nunique'}). \
    sort_values('count', ascending=False)
print(users.head())

df_final = df_final.rename(columns={'NamesBG': 'Име', 'used_on': 'Използван на', 'voucher_name': 'Име на ваучера',
                                    'Position': 'Позиция', 'Direction': 'Дирекция', 'Department': 'Отдел',
                                    'Team': 'Екип', 'count': 'Брой'})
df_final = df_final.drop('user_id', 1)
df_final = df_final.drop('domain_user', 1)

d['used_on'] = d.index
d = d[['used_on', 'count', 'user_id', 'voucher_name']]
d = d.rename(columns={'used_on': 'Номер', 'count': 'Брой получени ваучери', 'user_id': 'Брой служители, дали ваучер',
                      'voucher_name': 'Брой ваучери по вид'})

shops['team'] = shops.index
shops = shops[['team', 'count', 'user_id', 'voucher_name']]
shops = shops.rename(columns={'count': 'Брой подарени ваучери', 'user_id': 'Брой служители, дали ваучер',
                              'voucher_name': 'Брой ваучери по вид', 'team': 'Екип'})

users['NamesBG'] = users.index
users = users[['NamesBG', 'count', 'used_on', 'voucher_name']]
users = users.rename(columns={'count': 'Брой подарени ваучери', 'NamesBG': 'Служител',
                              'voucher_name': 'Брой ваучери по вид', 'used_on': 'Брой номера, получили ваучер'})

filename = "vouchers_report.xlsx"
writer = pd.ExcelWriter(filename,
                        engine='xlsxwriter',
                        date_format='dd.mm.yyyy',
                        datetime_format='dd.mm.yyyy hh:mm:ss')
df_final.to_excel(writer, sheet_name='All data', startrow=1, header=False, index=False, freeze_panes=(1, 0))
workbook = writer.book
worksheet = writer.sheets['All data']
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
for col_num, value in enumerate(df_final.columns.values):
    worksheet.write(0, col_num, value, header_format)
for idx, col in enumerate(df_final):  # loop through all columns
    series = df_final[col]
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
    )) + 2  # adding a little extra space
    # print(f"Series name: {str(series.name)}, length: {len(str(series.name))}")
    worksheet.set_column(idx, idx, max_len)  # set column width

# writer.save()

d.to_excel(writer, sheet_name='Ваучери по номера', startrow=1, header=False, index=False, freeze_panes=(1, 0))
workbook = writer.book
worksheet = writer.sheets['Ваучери по номера']
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
for col_num, value in enumerate(d.columns.values):
    worksheet.write(0, col_num, value, header_format)
for idx, col in enumerate(d):  # loop through all columns
    series = d[col]
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
    )) + 2  # adding a little extra space
    # print(f"Series name: {str(series.name)}, length: {len(str(series.name))}")
    worksheet.set_column(idx, idx, max_len)  # set column width

shops.to_excel(writer, sheet_name='Ваучери по екипи', startrow=1, header=False, index=False, freeze_panes=(1, 0))
workbook = writer.book
worksheet = writer.sheets['Ваучери по екипи']
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
for col_num, value in enumerate(shops.columns.values):
    worksheet.write(0, col_num, value, header_format)
for idx, col in enumerate(shops):  # loop through all columns
    series = shops[col]
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
    )) + 2  # adding a little extra space
    # print(f"Series name: {str(series.name)}, length: {len(str(series.name))}")
    worksheet.set_column(idx, idx, max_len)  # set column width

users.to_excel(writer, sheet_name='Ваучери по служители', startrow=1, header=False, index=False, freeze_panes=(1, 0))
workbook = writer.book
worksheet = writer.sheets['Ваучери по служители']
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
for col_num, value in enumerate(users.columns.values):
    worksheet.write(0, col_num, value, header_format)
for idx, col in enumerate(users):  # loop through all columns
    series = users[col]
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
    )) + 2  # adding a little extra space
    # print(f"Series name: {str(series.name)}, length: {len(str(series.name))}")
    worksheet.set_column(idx, idx, max_len)  # set column width

z = z.rename(columns={'months': 'Месец', 'NamesBG': 'Име', 'used_on': 'Използван на', 'voucher_name': 'Име на ваучера',
                      'Position': 'Позиция', 'Direction': 'Дирекция', 'Department': 'Отдел',
                      'Team': 'Екип'})
z = z.drop('user_id', 1)
z = z.drop('domain_user', 1)
print(z.head())

z.to_excel(writer, sheet_name='2 мес. активация', startrow=2, header=False, index=False, freeze_panes=(1, 0))
workbook = writer.book
worksheet = writer.sheets['2 мес. активация']

a1 = workbook.add_format({'bold': True, 'fg_color': '#98FB98', 'align': 'center'})
worksheet.merge_range('A1:D1',
                      "Активирани 2 поредни месеца един и същи ваучер от различни служители на един и същи номер",
                      a1)

header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
for col_num, value in enumerate(z.columns.values):
    worksheet.write(1, col_num, value, header_format)
for idx, col in enumerate(z):  # loop through all columns
    series = z[col]
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
    )) + 2  # adding a little extra space
    # print(f"Series name: {str(series.name)}, length: {len(str(series.name))}")
    worksheet.set_column(idx, idx, max_len)  # set column width
# worksheet.conditional_format('C3:C10000', {'type': '3_color_scale'})


writer.save()

send_mail.mail_send('Monthly vouchers report', to=['fraudinfo@a1.bg', 'n.iliykova@a1.bg', 'vn_mladenova@a1.bg'],
                    cc='', bcc=['a.zarkov@a1.bg', 'stefan.stoev@a1.bg'], email_text="<h3>Vouchers report for last 3 months</h3>",
                    att=['vouchers_report.xlsx'])
