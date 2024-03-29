import pandas as pd
import re
from pandas import ExcelWriter


print("——"*30)
print("Start!")
# Input File Name

filename=input("Input File Name: ")
filename2=input("Previous File: ")

print("Open Excel!")

file = pd.read_excel(filename)
file2 = pd.read_excel(filename2)


# Change Column Name
def change_column_name(filename, df):
    file_creater = filename[filename.find('，')+1:filename.find('.xlsx')]
    
    new_colname = []
    for col in df.columns:
        new_colname.append(str(col)+'_' + str(file_creater))
    df.columns = new_colname
    
    return df

def change_column_name2(filename, df, feature):
    new_colname = []
    for col in df.columns:
        new_colname.append(str(col)+'_' + str(feature))
    df.columns = new_colname
    
    return df

try:
    file = change_column_name(filename,file)
    prev_file = change_column_name(filename2,file2)
except:
    file = change_column_name2(filename, file, 'new')
    prev_file = change_column_name2(filename2, file2, 'old')

# Merge 2 Texts
def df_merge(df1, df2):
    whole = pd.merge(df1, df2,
                     left_on=list(df1.columns[2:5]),
                     right_on=list(df2.columns[2:5]),
                     how='outer')
    return whole

whole = df_merge(file,prev_file)

# Filter Normal Status if it has status
print("Check the differences between two charts!")
try:
    status = file.columns[file.columns.str.contains('状')][0]
    whole = whole[(whole[status] != '修改') & (whole[status] != '新增')]
except:
    pass

def find_column_new(string):
    lang = str(string)
    try:
        return file.columns[file.columns.str.contains(lang)][0]
    except:
        pass
    
def find_column_original(string):
    lang = str(string)
    try:
        return file2.columns[file2.columns.str.contains(lang)][0]
    except:
        pass
        
# To Change Column Sequence
col_list = []

def append_to_list(column_name):
    try:
        return col_list.append(column_name)
    except:
        pass

new_jianti = find_column_new('简体')
original_jianti = find_column_original('简体')
append_to_list(new_jianti)
append_to_list(original_jianti)

new_fanti = find_column_new('繁體')
original_fanti = find_column_original('繁體')
append_to_list(new_fanti)
append_to_list(original_fanti)

new_english = find_column_new('En')
original_english = find_column_original('En')
append_to_list(new_english)
append_to_list(original_english)

new_korean = find_column_new('한국어')
original_korean = find_column_original('한국어')
append_to_list(new_korean)
append_to_list(original_korean)

new_thai = find_column_new('ภาษา')
original_thai = find_column_original('ภาษา')
append_to_list(new_thai)
append_to_list(original_thai)

new_japanese = find_column_new('日本')
original_japanese = find_column_original('日本')
append_to_list(new_japanese)
append_to_list(original_japanese)

new_indonesian = find_column_new('Indon')
original_indonesian = find_column_original('Indon')
append_to_list(new_indonesian)
append_to_list(original_indonesian)

new_russian = find_column_new('Русс')
original_russian = find_column_original('Русс')
append_to_list(new_russian)
append_to_list(original_russian)    
    

def check_difference(df, new_lang, original_lang):
    df = df[df[new_lang] != df[original_lang]]
    
    return df

def change_col_sequence(changed_col_list, lang_column):
    try:
        changed_col_list.pop(changed_col_list.index(lang_column))
        changed_col_list.insert(0, lang_column)
        
        return changed_col_list
    except:
        pass
    
def changed_column(changed_col_list):
    try:
        return list(whole.columns[:5]) + changed_col_list
    except:
        pass
    

# Jianti
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_jianti)
changed_col_list = change_col_sequence(changed_col_list, new_jianti)

new_col_seq = changed_column(changed_col_list)
jianti = check_difference(whole, new_jianti, original_jianti)
jianti = jianti[new_col_seq]

# Fanti
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_fanti)
changed_col_list = change_col_sequence(changed_col_list, new_fanti)

new_col_seq = changed_column(changed_col_list)
fanti = check_difference(whole, new_fanti , original_fanti)[new_col_seq]

# English
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_english)
changed_col_list = change_col_sequence(changed_col_list, new_english)

new_col_seq = changed_column(changed_col_list)
english = check_difference(whole, new_english, original_english)[new_col_seq]

# Thai
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_thai)
changed_col_list = change_col_sequence(changed_col_list, new_thai)

new_col_seq = changed_column(changed_col_list)
thai = check_difference(whole, new_thai, original_thai)[new_col_seq]

# Korean
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_korean)
changed_col_list = change_col_sequence(changed_col_list, new_korean)

new_col_seq = changed_column(changed_col_list)
korean = check_difference(whole, new_korean, original_korean)[new_col_seq]

# Japanese
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_japanese)
changed_col_list = change_col_sequence(changed_col_list, new_japanese)

new_col_seq = changed_column(changed_col_list)
japanese = check_difference(whole, new_japanese, original_japanese)[new_col_seq]

# Indonesian
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_indonesian)
changed_col_list = change_col_sequence(changed_col_list, new_indonesian)

new_col_seq = changed_column(changed_col_list)
indonesian = check_difference(whole, new_indonesian, original_indonesian)[new_col_seq]

# Russian
changed_col_list = col_list.copy()
changed_col_list = change_col_sequence(changed_col_list, original_russian)
changed_col_list = change_col_sequence(changed_col_list, new_russian)

new_col_seq = changed_column(changed_col_list)
russian = check_difference(whole, new_russian, original_russian)[new_col_seq]


# Save to Different Sheet
def save_sheet(df, writer, sheet_name):
    try:
        return df.to_excel(writer, sheet_name)
    except:
        pass
    
with ExcelWriter(r'E:\desktop\1\Avid.ly\Harvest\words\z总表\check.xlsx') as writer:
    save_sheet(jianti, writer, '简体不同')
    save_sheet(fanti, writer, '繁体不同')
    save_sheet(english, writer, '英语不同')
    save_sheet(thai, writer, '泰语不同')
    save_sheet(korean, writer, '韩语不同')
    save_sheet(japanese, writer, '日语不同')
    save_sheet(indonesian, writer, '印尼语不同')
    save_sheet(russian, writer, '俄语不同')
    writer.save()

print("Check.xlsx has been created!\n")







