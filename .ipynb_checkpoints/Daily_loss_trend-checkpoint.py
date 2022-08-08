import pandas as pd
from pandas import ExcelWriter
filename=input("Input File Name or Drag Your File on here: ")
file = pd.read_excel(filename)

# Split texts to make a Data Frame
splited_data = file['新增人数'].str.split(' ',expand=True )

# All Index Rows Which is 步骤
buzhou = splited_data[splited_data[0]=='步骤'].index

datelist = []
pidlist = []
# NewU(2n)have 36 rows, OldU(2n-1) have 25 rows
for graph_num in range(len(buzhou)):
    if graph_num%2 == 0:
        # PID is 2rows Above the 步骤
        pid = splited_data.iloc[buzhou[graph_num]-2][0][1::]
        date = splited_data.iloc[buzhou[graph_num]-2][1]        
        # One for val, another for percentage
        pidlist.append(pid)
        pidlist.append(pid)
        datelist.append(date)
        datelist.append(date)

# New & Old User Data
datalist = []
# NewU(2n)have 36 rows, OldU(2n-1) have 25 rows
for graph_num in range(len(buzhou)):
    rows = buzhou[graph_num]
    if graph_num%2 == 0:
        new = splited_data.iloc[rows:rows+36].T
        datalist.append(new)
    else:
        old = splited_data.iloc[rows:rows+25].T
        datalist.append(old)

datalist = []
for graph_num in range(len(buzhou)):
# NewU(2n)have 36 rows, OldU(2n-1) have 25 rows
    start_row = buzhou[graph_num]
    if graph_num%2 == 0:
        # New Users 36 rows
        end_row = buzhou[graph_num]+36
    else:
        # Old Users 25 rows
        end_row = buzhou[graph_num]+25

    df = splited_data.iloc[start_row:end_row]
    # Reset Index Numbers
    df = df.reset_index()[[0,1,2]]
    # Transpose
    df = df.T
    # insert date at the first column
    df.insert(loc=0,column='date',value=datelist[graph_num])
    # insert pid at the first column
    df.insert(loc=0,column='pid',value=pidlist[graph_num])
    datalist.append(df)

newuser600167 = pd.DataFrame()
newuser_percent600167 = pd.DataFrame()
olduser600167 = pd.DataFrame()
olduser_percent600167 = pd.DataFrame()

newuser600168 = pd.DataFrame()
newuser_percent600168 = pd.DataFrame()
olduser600168 = pd.DataFrame()
olduser_percent600168 = pd.DataFrame()

# Findout each data by the sequence
newuserlist = [i*4  for i in range(10)]
newuser_percentlist = [i*4+1  for i in range(10)]
olduserlist = [i*4+2  for i in range(10)]
olduser_percentlist = [i*4+3  for i in range(10)]

def into_dataframe(dataframe, dataframe2, datalist_num):
    dataframe[num] = datalist[num].iloc[1]
    dataframe2[num] = datalist[num].iloc[2]

datalist_len_range = range(len(datalist))

# Check and put the data into each list
for num in datalist_len_range:
    if num in newuserlist:
        into_dataframe(newuser600167, newuser_percent600167, num)
    elif num in newuser_percentlist:
        into_dataframe(olduser600167, olduser_percent600167, num)
    elif num in olduserlist:
        into_dataframe(newuser600168, newuser_percent600168, num)
    elif num in olduser_percentlist:
        into_dataframe(olduser600168, olduser_percent600168, num)

def change_form(dataframe):
    dataframe = dataframe.T.sort_values(by='date',ascending=False)
    dataframe = dataframe.drop(columns=0)
    return dataframe
    
newuser600167 = change_form(newuser600167)
newuser_percent600167 = change_form(newuser_percent600167)
olduser600167 = change_form(olduser600167)
olduser_percent600167 = change_form(olduser_percent600167)
newuser600168 = change_form(newuser600168)
newuser_percent600168 = change_form(newuser_percent600168)
olduser600168 = change_form(olduser600168)
olduser_percent600168 = change_form(olduser_percent600168)
    
def change_columns(column_list):
    column_list[0] = 'PID'
    column_list[1] = 'date'
    column_list.remove('步骤')
    return column_list

newuser_collist = datalist[0].iloc[0].to_list()
newuser_collist = change_columns(newuser_collist)
olduser_collist = datalist[1].iloc[0].to_list()
olduser_collist = change_columns(olduser_collist)

newuser600167.columns = newuser_collist
newuser_percent600167.columns = newuser_collist
newuser600168.columns = newuser_collist
newuser_percent600168.columns = newuser_collist

olduser600167.columns = olduser_collist
olduser_percent600167.columns = olduser_collist
olduser600168.columns = olduser_collist
olduser_percent600168.columns = olduser_collist

# To excel, put in one sheet
writer = pd.ExcelWriter(location,engine='xlsxwriter')
workbook=writer.book
worksheet=workbook.add_worksheet('Sheet1')
writer.sheets['Sheet1'] = worksheet

newuser600167.to_excel(writer,sheet_name = 'Sheet1', startrow = 0 , startcol = 0)

startrow = len(newuser600167)+2
newuser_percent600167.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow, startcol = 0)

startrow = startrow + len(newuser_percent600167) + 2
olduser600167.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow, startcol = 0)

startrow = startrow + len(olduser600167) + 2
olduser_percent600167.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow , startcol = 0)

startrow = startrow + len(olduser_percent600167) + 2
newuser600168.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow, startcol = 0)

startrow = startrow + len(newuser600168) + 2
newuser_percent600168.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow, startcol = 0)

startrow = startrow + len(newuser_percent600168) + 2
olduser600168.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow, startcol = 0)

startrow = startrow + len(olduser600168) + 2
olduser_percent600168.to_excel(writer,sheet_name = 'Sheet1', startrow = startrow, startcol = 0)

writer.save()

print("每日流失.xlsx has been created!")

# To excel, put in different sheets
# with ExcelWriter(r'E:\desktop\Download\每日流失.xlsx') as writer:
#     newuser600167.to_excel(writer,'600167NAU')
#     newuser_percent600167.to_excel(writer,'600167NAU%')
#     olduser600167.to_excel(writer,'600167DOU')
#     olduser_percent600167.to_excel(writer,'600167DOU%')
#     newuser600168.to_excel(writer,'600168NAU')
#     newuser_percent600168.to_excel(writer,'600168NAU%')
#     olduser600168.to_excel(writer,'600168DOU')
#     olduser_percent600168.to_excel(writer,'600168DOU%')
# writer.save()