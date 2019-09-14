import pandas as pd

#ブック名入力
tdname=input('testdataName?')
edname=input('editordataName?')

#読み込んだブックの同じテストのデータをDataFrameに格納
td=pd.read_excel(tdname,header=1,sheet_name=0)
ed=pd.read_excel(edname,header=1,sheet_name=0)

#テストデータのラベルの部分をリストに変換
tdlabel=td.columns.values
tdlabel=tdlabel.tolist()

l=len(tdlabel)

##定義不要かもしれない
add=pd.DataFrame()
result=pd.DataFrame()

##エディタデータ成形
##rangeの引数は変数(input)にする必要があるかもしれない
for i in range(l):
    add = ed.loc[:,ed.columns.str.match(tdlabel[i])]
    result = pd.concat([result,add],axis=1)

tf=pd.DataFrame()
tf=td==result

#出力ブックの名前指定
outname=input('outputName?')

#各データを1ブック２シートに出力
with pd.ExcelWriter(outname) as writer:
    tf.to_excel(writer,sheet_name='TrueFalse')
    td.to_excel(writer,sheet_name='TestData')
    result.to_excel(writer,sheet_name='EditorData')
    

