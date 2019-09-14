#!/usr/bin/env python
# coding: utf-8

# In[6]:


import os
import glob

##C:\Users\founder3745\Desktop\test

repeatflag=0

while repeatflag==0:

  path=input("path\n")
  os.chdir(path)

  jpgname=path+"/*.jpg"
  jpgname

  flist=glob.glob(jpgname)

  print('変更前')
  print(flist)

  i=1

  for file in flist:
    os.rename(file, '00' + str(i) + '.JPG')
    i+=1
      
      
  list = glob.glob(jpgname)
  print('変更後')
  print(list)

  repeatflag=int(input("0:repeat 1:exit "))

# In[ ]:




