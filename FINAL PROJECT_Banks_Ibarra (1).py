#!/usr/bin/env python
# coding: utf-8

# In[574]:


import os
cwd = os.getcwd()
print(cwd)


# In[576]:


os.chdir('/Users/user/desktop')


# In[577]:


import xlrd


# In[578]:


loc = ("RNAseqdata.xls") #this function imports our data and allows to manipulate in the following function


# In[579]:


wb = xlrd.open_workbook(loc)


# In[580]:


sheet = wb.sheet_by_index(0)


# In[581]:


sheet.cell_value(1,4)


# In[582]:


print (sheet.nrows)


# In[583]:


import pandas as pd
xl = pd.ExcelFile("RNAseqdata.xls")


# In[584]:


cols = ['NAME', 'SCORE']
cols 
df = xl.parse("RNAseq")
df.drop('DESCRIPTION', axis =1, inplace = True)
df.drop('GENE_SYMBOL', axis =1, inplace = True)
df.drop('GENE_TITLE', axis =1, inplace = True)
df


# In[585]:


df = df.sort_values(by=['SCORE'])


# In[586]:


df #this function arraged our list to output top 5 genes with the highest score and the top 5 genes with the lowest score 


# In[587]:


df_high = df[df['SCORE'] >= 10] #based upon the literature 


# In[588]:


print (df_high [['NAME', 'SCORE']])


# In[589]:


import matplotlib.pyplot as plt 
import numpy as np 
from matplotlib.ticker import StrMethodFormatter


# In[590]:


df_high.plot(x='NAME', y='SCORE')


# In[591]:


df_low = df[df['SCORE'] <= -1] 


# In[592]:


print (df_low [['NAME', 'SCORE']])


# In[593]:


df_low.plot(x='NAME', y='SCORE')


# In[594]:


import pandas as pd
xl = pd.ExcelFile("RNAseqdata.xls")


# In[595]:


cols = ['Symbol']
df = xl.parse("UPR")
df


# In[596]:


df_upr = xl.parse("UPR")


# In[597]:


print (df_upr)


# In[598]:


from pandas import ExcelWriter
writer = pd.ExcelWriter('UpregulatedMatch.xlsx')
df_upr.to_excel(writer,'one', startrow=0 , startcol=0)
df_high.to_excel(writer,'one', startrow=0, startcol=2)
writer.save()


# In[599]:


import pandas as pd
df = pd.read_excel('UpregulatedMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[600]:


from pandas import ExcelWriter

writer = ExcelWriter('UpGraph.xlsx')
result.to_excel(writer,'one',index=False)
writer.save()


# In[601]:


import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.animation as animation
from IPython.display import HTML


# In[602]:


df = pd.read_excel('UpGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'UPR vs. RNASeq Matching Upregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[603]:


from pandas import ExcelWriter
writer = pd.ExcelWriter('DownregulatedMatch.xlsx')
df_upr.to_excel(writer,'one', startrow=0 , startcol=0)
df_low.to_excel(writer,'one', startrow=0, startcol=2)
writer.save()


# In[604]:


import pandas as pd
df = pd.read_excel('DownregulatedMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[605]:


from pandas import ExcelWriter

writer = ExcelWriter('DownGraph.xlsx')
result.to_excel(writer,'one',index=False)
writer.save()


# In[606]:


df = pd.read_excel('DownGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'UPR vs. RNASeq Matching Downregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[607]:


cols = ['Symbol']
df = xl.parse("Gly")
df


# In[608]:


df_Gly = xl.parse("Gly")


# In[609]:


print (df_Gly)


# In[610]:


writer = pd.ExcelWriter('GlyUpMatch.xlsx')
df_Gly.to_excel(writer,'Sheet1', startrow=0 , startcol=0)
df_high.to_excel(writer,'Sheet1', startrow=0, startcol=2)
writer.save()


# In[611]:


df = pd.read_excel('GlyUpMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[612]:


writer = ExcelWriter('GlyUpGraph.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()


# In[613]:


df = pd.read_excel('GlyUpGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'Gly vs. RNASeq Matching Upregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[614]:


writer = pd.ExcelWriter('GlyDownMatch.xlsx')
df_Gly.to_excel(writer,'Sheet1', startrow=0 , startcol=0)
df_low.to_excel(writer,'Sheet1', startrow=0, startcol=2)
writer.save()


# In[615]:


df = pd.read_excel('GlyDownMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[616]:


writer = ExcelWriter('GlyDownGraph.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()


# In[617]:


df = pd.read_excel('GlyDownGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'Gly vs. RNASeq Matching Downregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




