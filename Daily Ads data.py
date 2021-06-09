import pandas as pd
import os
import sqlite3
import sys

con = sqlite3.connect(r"E:\desktop\1\Avid.ly\Harvest\Data\g广告相关\每日广告\ht_daily_ads.sqlite")
print("——"*30)
print("Change the special character in the file name '—' to '_' first!")

filename=input("Input File Name:")
print("File Name: ",filename)

file = pd.read_excel(filename)

# From 20xx to .xlsx
date = pd.to_datetime(filename[filename.find('2'):filename.find('.xlsx')])

# Check if the data is already in the DB 
SQL="""
select * from ad_unit_daily
where 日期 == '{}'
limit 1
""".format(date.strftime('%Y-%m-%d'))
checking = pd.read_sql_query(SQL,con)

if len(checking) ==1:
    print("Data Already exists in DB")
    
elif len(checking) ==0:
    # Read Excel and save it to DB
    open_file_name = 'E:\\desktop\\Download' + filename[filename.find('600167')-1::]
    print(open_file_name)
    print("Put into DB")
    print("——"*30)
    file = file[['product_id', '日期', 'dau', '广告收益', '广告位', '广告入口点击', '展示次数', '展示人数']]                
    file.to_sql("ad_unit_daily", con, if_exists="append",index=False)
    
    # Change the type of date
    SQL="""
    select * from ad_unit_daily
    where 日期 == '{}'
    """.format(str(date))
    date_data = pd.read_sql_query(SQL,con)
    date_data.日期 = date_data.日期.str.split(expand=True)[0]
    
    # Delete the Original Data
    cur = con.cursor()
    SQL="""
    Delete from ad_unit_daily
    where 日期 == '{}'
    """.format(str(date))

    cur.execute(SQL)
    con.commit()
    cur.close()

    # Put the new date type data into DB
    date_data.to_sql("ad_unit_daily", con, if_exists="append",index=False)
    
con.close()    

# Move File
move_dir = r'E:\desktop\1\Avid.ly\Harvest\Data\g广告相关\每日广告' + filename[filename.find('600167')-1::]
os.replace(filename, move_dir)

print("——"*30)
print("Finished!")