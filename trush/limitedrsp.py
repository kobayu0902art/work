#finger使い切った方が強い
#もしくはrwin=2point,swin=3point,pwin=5point
import itertools

n=int(input('n回じゃんけん\n'))
finger=int(input('number of finger\n'))

combidict={}
for p in range((finger//5)+1):
    for s in range(((finger-p*5)//2)+1):
        if p>n or s>n:
            continue
        r=n-p-s
        total=p*5+s*2
        #print(f'Rock:{r},Scissors:{s},Paper:{p},total:{total}')
        #ここまでbasic
        
        #ここから追加
        dictkey=str(r)+','+str(s)+','+str(p)
        combidict[dictkey]=total

#考えられる組み合わせすべて
combidict_sorted=sorted(combidict.items(), key=lambda x:x[1])
#print(combidict_sorted)

#finger使い切る組み合わせ
keys = [k for k,v in combidict.items() if v == finger]
#print(f'len:{len(combidict)}')
print(f'keys:{keys},len:{len(keys)}')

#ここから手の組み合わせ(最後のr,s,pのみ)
charlist=[]
for n in range(r):
    charlist.append('r')
for n in range(s):
    charlist.append('s')
for n in range(p):
    charlist.append('p')
#print(charlist)
#print(r,s,p)
combi=set(list(itertools.permutations(charlist)))
print(f'combi:\n{combi}\nlen:{len(combi)}')