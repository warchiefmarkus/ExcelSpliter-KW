# -*- coding: utf-8 -*-

from pandas import ExcelWriter
from pandas import DataFrame
import pandas as pd
from functools import reduce
import numpy as np

def splitBy(iterable, where):
    def splitter(acc, item, where=where):
        if item == where:
            acc.append([])
        else:
            acc[-1].append(item)
        return acc
    return reduce(splitter, iterable, [[]])


# GET SHEET TO DF
xl = pd.ExcelFile("C:/Users/user/Desktop/file.xlsx")
df = xl.parse("sheet1")

#%%

# CONCAT COLUMNS
def concatCol(x):
    sent_str = ""
    for index in range(len(splitBy(df.columns.tolist(),'ИГРОВОЙ ТРАНСПОРТ')[1])):
        if x[index]!="":
            if splitBy(df.columns.tolist(),'ИГРОВОЙ ТРАНСПОРТ')[1][index]=='ПРИМЕЧАНИЯ':
                sent_str += x[index]+"\n"
            else:
                sent_str += splitBy(df.columns.tolist(),'ИГРОВОЙ ТРАНСПОРТ')[1][index]+" : "+x[index]+"\n"
    return sent_str

df['ПРИМЕЧАНИЯ'] = df[splitBy(df.columns.tolist(),'ИГРОВОЙ ТРАНСПОРТ')[1]].fillna("").apply(concatCol, axis=1)
df = df.drop(splitBy(df.columns.tolist(),'ПРИМЕЧАНИЯ')[1],axis=1)

#%%

# SPLIT BY GROUPS BASED ON KPP COLUMN
gb = df.groupby(["КПП"])
splited = [gb.get_group(x) for x in gb.groups]
#%%

# EXPORT
writer = pd.ExcelWriter('C:/Users/leo/Desktop/test/book.xlsx', engine='xlsxwriter')
for i in splited:
    k = i.drop(splitBy(df.columns.tolist(),'СЕР.')[0],axis=1) 
#5 LINES
#    for f in range(0,5):
#        df1 = pd.DataFrame([[""] * len(k.columns)], columns=k.columns)
#        k = df1.append(k, ignore_index=True)
        
    k.to_excel(writer, sheet_name=(i['КПП'].iloc[0]).replace("/","-"),index=False)
        
    workbook  = writer.book
    worksheet = writer.sheets[(i['КПП'].iloc[0]).replace("/","-")]
    format1 = workbook.add_format()
    format2 = workbook.add_format({'num_format': 'hh:mm:ss'})
    format1.set_text_wrap()
    format1.set_font_name('Calibri')
    format2.set_font_name('Calibri')
    format1.set_font_size('11')
    format2.set_font_size('11')
        
    worksheet.set_column('A:A', 5.29, format1)
    worksheet.set_column('B:B', 5.29, format1)
    worksheet.set_column('C:C', 5.29, format1)
    worksheet.set_column('D:D', 5.29, format1)
    worksheet.set_column('E:E', 8.50, format2)
    worksheet.set_column('F:F', 5.30, format1)
    worksheet.set_column('G:G', 16.30, format1)
    worksheet.set_column('H:H', 15, format1)
    worksheet.set_column('I:I', 32.30, format1)
    worksheet.set_column('J:J', 0, format1)
    worksheet.set_column('K:K', 16, format1)
    worksheet.set_column('L:L', 16.30, format1)
    worksheet.set_column('M:M', 23, format1)
    worksheet.set_column('N:N', 19, format1)
    worksheet.set_column('O:O', 34, format1)
    
    # Add a header format.
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'border': 1})

#    # Write the column headers with the defined format.
#    for col_num, value in enumerate(i):
#        worksheet.write(0, col_num + 1, value, header_format)
    
writer.save()
writer.close()
