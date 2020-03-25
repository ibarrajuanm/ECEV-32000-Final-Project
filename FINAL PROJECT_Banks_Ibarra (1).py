#!/usr/bin/env python
# coding: utf-8

# In[512]:


import os
cwd = os.getcwd()
print(cwd)


# In[513]:


os.chdir('/Users/user/desktop')


# In[514]:


import xlrd


# In[515]:


loc = ("RNAseqdata.xls") #this function imports our data and allows to manipulate in the following function


# In[516]:


wb = xlrd.open_workbook(loc)


# In[517]:


sheet = wb.sheet_by_index(0)


# In[518]:


sheet.cell_value(1,4)


# In[519]:


print (sheet.nrows)


# In[522]:


import pandas as pd
xl = pd.ExcelFile("RNAseqdata.xls")


# In[525]:


cols = ['NAME', 'SCORE']
cols 
df = xl.parse("RNAseq")
df.drop('DESCRIPTION', axis =1, inplace = True)
df.drop('GENE_SYMBOL', axis =1, inplace = True)
df.drop('GENE_TITLE', axis =1, inplace = True)
df


# In[527]:


df = df.sort_values(by=['SCORE'])


# In[528]:


df #this function arraged our list to output top 5 genes with the highest score and the top 5 genes with the lowest score 


# In[529]:


df_high = df[df['SCORE'] >= 10] #based upon the literature 


# In[530]:


print (df_high [['NAME', 'SCORE']])


# In[531]:


import matplotlib.pyplot as plt 
import numpy as np 
from matplotlib.ticker import StrMethodFormatter


# In[532]:


df_high.plot(x='NAME', y='SCORE')


# In[533]:


df_low = df[df['SCORE'] <= -1] 


# In[534]:


print (df_low [['NAME', 'SCORE']])


# In[535]:


df_low.plot(x='NAME', y='SCORE')


# In[536]:


import pandas as pd
xl = pd.ExcelFile("RNAseqdata.xls")


# In[537]:


cols = ['Symbol']
df = xl.parse("UPR")
df


# In[539]:


df_upr = xl.parse("UPR")


# In[540]:


print (df_upr)


# In[543]:


from pandas import ExcelWriter
writer = pd.ExcelWriter('UpregulatedMatch.xlsx')
df_upr.to_excel(writer,'one', startrow=0 , startcol=0)
df_high.to_excel(writer,'one', startrow=0, startcol=2)
writer.save()


# In[544]:


import pandas as pd
df = pd.read_excel('UpregulatedMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[545]:


from pandas import ExcelWriter

writer = ExcelWriter('UpGraph.xlsx')
result.to_excel(writer,'one',index=False)
writer.save()


# In[546]:


import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.animation as animation
from IPython.display import HTML


# In[570]:


df = pd.read_excel('UpGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'UPR vs. RNASeq Matching Upregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[548]:


from pandas import ExcelWriter
writer = pd.ExcelWriter('DownregulatedMatch.xlsx')
df_upr.to_excel(writer,'one', startrow=0 , startcol=0)
df_low.to_excel(writer,'one', startrow=0, startcol=2)
writer.save()


# In[549]:


import pandas as pd
df = pd.read_excel('DownregulatedMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[550]:


from pandas import ExcelWriter

writer = ExcelWriter('DownGraph.xlsx')
result.to_excel(writer,'one',index=False)
writer.save()


# In[571]:


df = pd.read_excel('DownGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'UPR vs. RNASeq Matching Downregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[553]:


cols = ['Symbol']
df = xl.parse("Gly")
df


# In[554]:


df_Gly = xl.parse("Gly")


# In[555]:


print (df_Gly)


# In[556]:


writer = pd.ExcelWriter('GlyUpMatch.xlsx')
df_Gly.to_excel(writer,'Sheet1', startrow=0 , startcol=0)
df_high.to_excel(writer,'Sheet1', startrow=0, startcol=2)
writer.save()


# In[557]:


df = pd.read_excel('GlyUpMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[558]:


writer = ExcelWriter('GlyUpGraph.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()


# In[572]:


df = pd.read_excel('GlyUpGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'Gly vs. RNASeq Matching Upregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[560]:


writer = pd.ExcelWriter('GlyDownMatch.xlsx')
df_Gly.to_excel(writer,'Sheet1', startrow=0 , startcol=0)
df_low.to_excel(writer,'Sheet1', startrow=0, startcol=2)
writer.save()


# In[561]:


df = pd.read_excel('GlyDownMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[562]:


writer = ExcelWriter('GlyDownGraph.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()


# In[573]:


df = pd.read_excel('GlyDownGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'Gly vs. RNASeq Matching Downregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[ ]:




