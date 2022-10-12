import cx_Oracle
import pandas as pd
import sys
from datetime import datetime, timedelta, date
import os

sys.path.insert(0, 'c:/xampp/htdocs/includes/')
import db_config

path = r"\\cassiopeia\Work\Finance\TreasuryOperation\BankCashAccount\Shops Database\Robot\Cash_differences\ДСК"


def get_incasodb(date_for):
    con = cx_Oracle.connect(db_config.user, db_config.pw, db_config.dsn, encoding="UTF-8", nencoding="UTF-8")
    # cur = con.cursor()
    dd = date_for.strftime('%Y-%m-%d')
    sap = r'\d{4}$'
    query = f"""
        select DATE_INSERT, 
            DETAIL1 || DETAIL2 as details, 
            AMOUNT, 
            to_number(regexp_substr(detail2, '{sap}')) as sap_code
        from MT940
        where trunc(MT940.DATE_INSERT) = trunc(to_date('{dd}', 'YYYY-MM-DD'))
        and FILTERED = 'Y'
        and (MT940.DETAIL1 like '%ИЗЛИШЪК%' or MT940.DETAIL1 like '%ДЕФИЦИТ%')    
        order by 1
    """
    # print(query)
    df = pd.read_sql(query, con)
    con.close()

    if len(df) == 0:
        print("No data for this date.")
        return 1
    else:
        print(f"Found {len(df)} rows.")
    print(df)
    try:
        filename = f'{path}\\{date_for.year}\\{date_for.month:02d}.{date_for.year}\\' \
                   f'{date_for.day:02d}.{date_for.month:02d}.{date_for.year}.xlsx'
        writer = pd.ExcelWriter(filename,
                                engine='xlsxwriter',
                                date_format='dd.mm.yyyy',
                                datetime_format='dd.mm.yyyy hh:mm:ss')
        df.to_excel(writer, sheet_name='Incaso', startrow=1, header=False, index=False, freeze_panes=(1, 0))
        workbook = writer.book
        worksheet = writer.sheets['Incaso']
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'top',
            'fg_color': '#F08080',
            'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
            )) + 2  # adding a little extra space
            # print(f"Series name: {str(series.name)}, length: {len(str(series.name))}")
            worksheet.set_column(idx, idx, max_len)  # set column width
        writer.save()
    except Exception as e:
        print("Error creating file!")
        print(e)
        return 1


def check_dir(dt):
    fdir = f'{path}\\{dt.year}\\{dt.month:02d}.{dt.year}'
    # print(fdir)
    try:
        dir_exists = os.path.isdir(fdir)
    except Exception as e:
        print("Error checking directory. Please investigate.")
        print(e)
        return 2
    # print(dir_exists)
    if not dir_exists:
        print(f"Directory {fdir} does not exist. Creating...")
        try:
            os.makedirs(fdir)
        except Exception as e:
            print("Error with directory creating. Please check.")
            print(e)
            return 1
    return 0


days_ago = 1
yesterday = date.today() - timedelta(days=days_ago)
datelist = pd.date_range(datetime.today() - timedelta(days=days_ago), periods=days_ago).tolist()
# print(datelist)
for dl in datelist:
    print(f"Checking date {dl.strftime('%Y-%m-%d')}")
    if check_dir(dl) > 0:
        break
    get_incasodb(dl)
