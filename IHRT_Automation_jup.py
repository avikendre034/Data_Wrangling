#!/usr/bin/env python
# coding: utf-8

# In[35]:


import pandas as pd
pd.set_option('display.max_columns', None)                                        # Unfolding hidden features if the cardinality is high      
pd.set_option('display.max_colwidth', None)                                       # Unfolding the max feature width for better clearity      
pd.set_option('display.max_rows', None)                                           # Unfolding hidden data points if the cardinality is high
pd.set_option('display.float_format', lambda x: '%.2f' % x)   
import numpy as np
from win10toast import ToastNotifier
from pathlib import Path
import os
import six
import appdirs
import packaging.requirements
import textwrap
import pyodbc


# In[109]:


# Locate examples files & create output directory
#output_dir = Path.cwd() / "Output"
ecl_file_path =r'C:\Users\kendrav\OneDrive - Ecolab\APAC_Tier 2_V2\Inventory Health Report Tier 2 - Transformation file ECL - P06 update (Chem & EQ combined).xlsx'
nalco_file_path =r'C:\Users\kendrav\OneDrive - Ecolab\APAC_Tier 2_V2\Inventory Health Report Tier 2 - Transformation File NCL - P06 update.xlsx'
#output_dir.mkdir(exist_ok=True)


# In[131]:


ecl_df=pd.read_excel(ecl_file_path,sheet_name='Pivot',header=4,usecols="A:AC")


# In[132]:


print("Ecolab File Data shape as :",ecl_df.shape)


# In[133]:


ncl_df = pd.read_excel(nalco_file_path,sheet_name='Pivot - NCL',header=4,usecols="A:AC")


# In[134]:


print("Nalco File Data shape as :",ncl_df.shape)


# In[135]:


combine_df = pd.concat([ecl_df,ncl_df],ignore_index=True)


# In[136]:


print("Combined Data shape as :",combine_df.shape)


# In[137]:


combine_df.replace(to_replace='(blank)',value=None,inplace=True)


# In[138]:


combine_df = combine_df.convert_dtypes()


# In[139]:


#combine_df.info()


# In[140]:


#combine_df.describe(include='all')


# In[ ]:





# In[141]:


combine_df.to_excel(r'\\global.ecolab.corp\apac\IN-Pune\GROUPS\Planning_Analytics\Projects\Avinash_Automation_NW\IHRT\Final_IHRT.xlsx',index=False)


# In[91]:


n = ToastNotifier()
n.show_toast("IHRT Automation Notification","Final File Generated in Network IHRT Folder !!!", duration=15)


# In[ ]:




