#import openpyxl as px

on=input("sample")
l = [int(i) for i in input().split()]

for i in range(len(l)):
    jj=1
    for j in range(l[i]):
        ii=int(i)
        ii+=1
        ii=str(i+1)
        ii=ii.zfill(2)
        jj=str(jj)
        jj=jj.zfill(3)
        print("sample"+on+"-"+ii+"-"+jj)
        jj=int(jj)
        jj+=1