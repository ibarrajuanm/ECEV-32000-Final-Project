#!/usr/bin/env python
# coding: utf-8

# In[156]:


#for better organization we have users import all modules that will be needed throughout the code 
import os

#This library allows to read and format Excel files
#Import RNAseqdata in excel format 
import xlrd 

#This library provides with multiple operations and data analysis 
#import working data which contains CALR vs. CALR mutant RNA seq. results and the lists of genes associated with the unfolded protein response (UPR) and Glycolysis (Gly) pathways obtained from Reactome database
import pandas as pd

#This method allows to export dataframe objects into excel files
#Write in excel the matching outcomes for upregulated and downregulated genes
from pandas import ExcelWriter

#This library allows to generate plots and other visualization graphs
#Plot downregulated and upregulated genes + matching genes
import matplotlib.pyplot as plt 

#This library provides with high-level mathematical functions and also provides with graphic functions
#Used along with matplot to generate graphs of upregulated and downregulated genes 
import numpy as np 

#This function allows to define and format ticks, and set scale
import matplotlib.ticker as ticker

#This function supplies a string that can be formatted using the format method and also provides a cleaner visualization of the axis values 
from matplotlib.ticker import StrMethodFormatter

#This function allows to display representation of graphs in the format of HTML
from IPython.display import HTML


# In[157]:


#set working directory and upload your data excel sheet of interest
cwd = os.getcwd()
os.chdir('/Users/user/desktop')
#User defines file name userinput = ("enter file name needed to read")
userinput = ("RNAseqdata.xls")
myfile = open(userinput)
#User selects threshold for significant upregulated and downregulated genes
Threshold_Upregulated = 20
Threshold_Downregulated = -20


# In[158]:


cols = ['NAME', 'SCORE']
df = xl.parse("RNAseq")
df.drop('DESCRIPTION', axis =1, inplace = True)
df.drop('GENE_SYMBOL', axis =1, inplace = True)
df.drop('GENE_TITLE', axis =1, inplace = True)
df


# In[159]:


#sort CALR vs. CALR mutant RNA Seq. genes by the given expression score from lowest to highest
df = df.sort_values(by=['SCORE'])
df


# In[160]:


cols = ['NAME', 'SCORE']
df = xl.parse("RNAseq")
df.drop('DESCRIPTION', axis =1, inplace = True)
df.drop('GENE_SYMBOL', axis =1, inplace = True)
df.drop('GENE_TITLE', axis =1, inplace = True)
df


# In[161]:


#sort CALR vs. CALR mutant RNA Seq. genes by the given expression score from lowest to highest
df = df.sort_values(by=['SCORE'])
df


# In[162]:


#set threshold and define range of genes thought to be significantly upregulated based on threshold (threshold value can be set based on the literature, in this case we use 10 as an arbitrary value)
#Anything above score 10 is considered significantly upregulated 
#Define upregulated genes data frame
df_Upregulated = df[df['SCORE'] >= Threshold_Upregulated] 
#list of significantly upregulated genes
print (df_Upregulated [['NAME', 'SCORE']])


# In[163]:


#Graphing of the significantly upregulated genes, this gives the user the ability to visualize the general trend of the results
#Graph significantly upregulated genes 
df_Upregulated.plot(x='NAME', y='SCORE', kind= 'bar', figsize=(20,15), fontsize =25)
plt.xlabel("Gene", labelpad=1, fontsize =25)
plt.ylabel("Expression Score", labelpad=1, fontsize =25)
plt.title("CALR Mutant vs. CALR WT RNAseq Significantly Upregulated Genes", y=1, fontsize =25)


# In[164]:


#set threshold and define range of genes thought to be significantly downregulated based on threshold (threshold value can be set based on the literature, in this case we use -5 as an arbitrary value)
#Anything above score -5 is considered significantly downregulated 
#define downregulated genes dataframe
df_downregulated = df[df['SCORE'] <= Threshold_Downregulated] 


# In[165]:


#list significantly downregulated genes based on threshold 
print (df_downregulated [['NAME', 'SCORE']])


# In[166]:


#Graph of significantly downregulated genes
df_downregulated.plot(x='NAME', y='SCORE', kind= 'bar', figsize=(45,20),fontsize =40)
plt.ylabel("Expression Score", labelpad=0.5, fontsize=40)
plt.title("CALR vs. CALR WT RNAseq Significantly Downregulated Genes", y=1, fontsize =40)
plt.xlabel("Gene", labelpad=1, fontsize =40)


# In[167]:


#Select UPR associated genes to be matched against significantly down or upregulated RNA seq. genes
#Discriminate columns: only information required from the UPR associated genes data is the gene "symbol"
cols = ['Symbol']
df = xl.parse("UPR")
df


# In[168]:


#Define UPR associated genes as dataframe
df_upr = xl.parse("UPR")
#check/list UPR data frame
print (df_upr)


# In[169]:


#Write UPR and RNA seq. upregulated genes in one excel file named UPRUpMatch.xlsx and save
writer = pd.ExcelWriter('UPRUpMatch.xlsx')
df_upr.to_excel(writer,'one', startrow=0 , startcol=0)
df_Upregulated.to_excel(writer,'one', startrow=0, startcol=2)
writer.save()


# In[170]:


#Graph matched genes (include gene name and expression score) 
df = pd.read_excel('UpGraph.xlsx', usecols =['NAME', 'SCORE'], fontsize= 30)
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'UPR vs. RNASeq Matching Upregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left', fontsize= 30)


# In[171]:


#Write UPR and RNA seq. downregulated genes in one excel file named UPRDownMatch.xlsx
writer = pd.ExcelWriter('UPRDownMatch.xlsx')
df_upr.to_excel(writer,'one', startrow=0 , startcol=0)
df_downregulated.to_excel(writer,'one', startrow=0, startcol=2)
writer.save()


# In[172]:


#Match UPR and downregulated genes based gene name
#Whichever genes are found in both or match UPR and RNA seq. downregulated columns will appear in list below with the corresponding expression score
df = pd.read_excel('UPRDownMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[173]:


#Import the chart above as output into a separate excel sheet named UPRDownGraph.xlsx
writer = ExcelWriter('UPRDownGraph.xlsx')
result.to_excel(writer,'one',index=False)
writer.save()


# In[174]:


#Generate graph of the matched UPR and downregulated RNA Seq. genes 
#Graph matched genes (include gene name and score)

df = pd.read_excel('UPRDownGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'UPR vs. RNASeq Matching Downregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[175]:


#Select Glycolysis (Gly) associated genes to be matched against significantly down or upregulated RNA seq. genes
#Discriminate columns: only information required from the Gly associated genes data is the gene "symbol"
cols = ['Symbol']
df = xl.parse("Gly")
df


# In[176]:


#Define Gly associated genes as dataframe
df_Gly = xl.parse("Gly")
#check/list Gly data frame
print (df_Gly)


# In[177]:


#Write Gly and RNA seq. upregulated genes in one excel file named GlyUpMatch.xlsx

writer = pd.ExcelWriter('GlyUpMatch.xlsx')
df_Gly.to_excel(writer,'Sheet1', startrow=0 , startcol=0)
df_Upregulated.to_excel(writer,'Sheet1', startrow=0, startcol=2)
writer.save()


# In[178]:


#Drop non-required colums/data from data imported in the GlyUpMatch.xlsx
#Match Gly and Upregulated genes based gene name
#Whichever genes are found in both or match Gly and RNA seq. upregulated columns will appear in list below with the corresponding expression score

df = pd.read_excel('GlyUpMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[179]:


#Import the above chart of matching genes as output into a seperate excel sheet named GlyUpGraph.xlsx
writer = ExcelWriter('GlyUpGraph.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()


# In[180]:


#Graph matched genes (include gene name and score) 
df = pd.read_excel('GlyUpGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'Gly vs. RNASeq Matching Upregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[181]:


#Write Gly and RNA seq. downregulated genes in one excel file named GlyDownMatch.xlsx
writer = pd.ExcelWriter('GlyDownMatch.xlsx')
df_Gly.to_excel(writer,'Sheet1', startrow=0 , startcol=0)
df_downregulated.to_excel(writer,'Sheet1', startrow=0, startcol=2)
writer.save()


# In[182]:


#Match Gly and downregulated genes based on gene name
#Whichever genes are found in both or match UPR and RNA seq. downregulated columns will appear in list below with the corresponding expression score
df = pd.read_excel('GlyDownMatch.xlsx')
df.drop(df.columns[df.columns.str.contains('Unnamed',case = False)],axis = 1, inplace = True)
result = df[df['NAME'].isin(list(df['Symbol']))]
result


# In[183]:


#Import the chart above as output into a separate excel sheet named GlyDownGraph.xlsx
writer = ExcelWriter('GlyDownGraph.xlsx')
result.to_excel(writer,'Sheet1',index=False)
writer.save()


# In[184]:


#Generate graph of the matched Gly and downregulated RNA Seq. genes 
#Graph matched genes (include gene name and score)

df = pd.read_excel('GlyDownGraph.xlsx', usecols =['NAME', 'SCORE'])
fig, ax = plt.subplots(figsize=(8, 4))
ax.barh(df['NAME'], df['SCORE'])
ax.text(0, 1.12, 'Gly vs. RNASeq Matching Downregulated Genes',
            transform=ax.transAxes, size=24, weight=600, ha='left')


# In[ ]:





# In[ ]:




