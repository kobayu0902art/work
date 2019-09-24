#!/usr/bin/env python
# coding: utf-8
import os
import pandas as pd
import openpyxl as px
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Color, PatternFill
import unicodedata
def is_japanese(string):
    for ch in string:
        name = unicodedata.name(ch) 
        if "CJK UNIFIED" in name \
        or "HIRAGANA" in name \
        or "KATAKANA" in name:
            return True
    return False
spec=r"sample"
evid=r"sample"
bs="\\"
xls=".xls"
spec+=bs
evid+=bs
assignbook=r"sample"
checkbook=r"sample"
wb = px.load_workbook(assignbook)
wb.save(checkbook)
ws=wb.active
ws.cell(row=1,column=2,value="日付")
ws.cell(row=1,column=3,value="仕様書エビデンスファイル数")
ws.cell(row=1,column=4,value="実際のエビデンスファイル数")
ws.cell(row=1,column=5,value="CD TrueFalse")
ws.cell(row=1,column=6,value="仕様書エビデンスファイル名")
ws.cell(row=1,column=7,value="実際のエビデンスファイル名")
worklist=pd.read_excel(assignbook)
worklist=worklist.values.tolist()
l=len(worklist)
for i in range(l):
    one=str(worklist[i])[2:9]
    two=str(worklist[i])[2:12]
    three=str(worklist[i])[2:16]
    finalspecpath=spec+one+bs+two+bs+three+xls
    finalevidpath=evid+one+bs+two+bs+three
    filelist=os.listdir(finalevidpath)
    fllen=len(filelist)
    specwb=pd.read_excel(finalspecpath, header=None) 
    evidarea=specwb.loc[20:60,2:40]
    evidarea.to_excel(r"sample")
    tempwb=px.load_workbook(r"sample")
    tempws=tempwb.active
    evidlist=list(range(1))
    for col_num in range(1,40):
        for row_num in range(1,42):
            if tempws.cell(row=row_num,column=col_num).value == None:
                tempws.cell(row=row_num,column=col_num,value="temp")
            jap=is_japanese(str(tempws.cell(row=row_num,column=col_num).value))
            if "sample" in str(tempws.cell(row=row_num,column=col_num).value) and jap == False:
                evidlist.append(tempws.cell(row=row_num,column=col_num).value)
    del evidlist[0]
    samplelen=len(evidlist)
    ws.cell(row=i+2,column=2,value=str(specwb.at[1,38]))
    ws.cell(row=i+2,column=3,value=samplelen)
    ws.cell(row=i+2,column=4,value=fllen)
    ipt=str(i+2)
    tf="=C"+ipt+"=D"+ipt
    ws.cell(row=i+2,column=5,value=tf)
    ws.cell(row=i+2,column=6,value=str(evidlist))
    ws.cell(row=i+2,column=7,value=str(filelist))
range="E2:E"+ipt
ws.conditional_formatting.add(range,
                              CellIsRule(operator='equal',formula=['FALSE'], 
                                         fill=PatternFill(start_color='FF0000', end_color='FF0000',fill_type='solid')))
os.remove(r"sample")
wb.save(checkbook)
