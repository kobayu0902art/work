import os
import glob
import re


primary=[
    r"path1",
    r"path2",
    r"path3",
    r"path4",
    r"path5"
]
secondary=[
    r"path6",
    r"path7"
]

pl=len(primary)
sl=len(secondary)

for q in range(pl):
    primary[q]+="\\"
    #print(primary[q])
    
for q in range(sl):
    secondary[q]+="\\"
    #print(secondary[q])


flag=0
xlsx=".xlsx"
xls=".xls"

def primary_search(pl,xlsx):
    for i in range(pl):
            os.chdir(primary[i])
            tempfilename=primary[i]+"**\\"+filename+xlsx
            l=glob.glob(tempfilename, recursive=True)
            if bool(l)==True:
                print(l) 
            if bool(l)==False:
                print("該当なし") 
                
def secondary_search(pl,xlsx):
    for i in range(pl):
            os.chdir(secondary[i])
            tempfilename=secondary[i]+"**\\"+filename+xlsx
            l=glob.glob(tempfilename, recursive=True)
            if bool(l)==True:
                print(l) 
            if bool(l)==False:
                print("該当なし")
    
while flag==0:
    filename=input("file name?\n")
    print("優先パスを検索します")
    print("xlsxを検索")
    primary_search(pl,xlsx)
    print("xlsを検索")
    primary_search(pl,xls)
    print("\n")
    exitflag=int(input("さらに過去も検索:0,exit:1\n"))
    if exitflag==1:
        break
    else:
        print("xlsxを検索")
        secondary_search(sl,xlsx)
        print("xlsを検索")
        secondary_search(sl,xls)
        print("\n")
        flag=int(input("0:continue\n1:exit\n"))    
input("end:enter")