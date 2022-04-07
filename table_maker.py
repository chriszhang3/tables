#!/usr/bin/env python
# coding: utf-8

# In[1]:


"""Convert a csv file to a table in a Microsoft Word document."""
import os
import pandas as pd
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT


# In[2]:


filename = "test" # Change this. Should be the name of the file without .csv
directory = "data"
df = pd.read_csv(os.path.join(directory, f"{filename}.csv"), index_col=0)
df.head()


# In[3]:


doc = docx.Document()

# add a table to the end and create a reference variable
# extra row is so we can add the header row
table = doc.add_table(df.shape[0]+1, df.shape[1])
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# add the header rows.
for j in range(df.shape[-1]):
    table.cell(0,j).text = df.columns[j]
    table.cell(0,j).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

# add the rest of the data frame
for i in range(df.shape[0]):
    for j in range(df.shape[-1]):
        table.cell(i+1,j).text = str(df.values[i,j])
        table.cell(i+1,j).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER


# In[4]:


doc.save(os.path.join(directory, f"{filename}.docx"))


# In[ ]:




