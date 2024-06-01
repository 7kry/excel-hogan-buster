#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xlwings as xw
import shutil
import os.path
import pandas as pd
import operator
import math
import itertools
import re
import tkinter
from tkinter import filedialog as tkfd
import sys


# In[2]:


# tkinterおまじない
root = tkinter.Tk()
root.attributes('-topmost', True)
root.withdraw()
root.lift()
root.focus_force()


# In[3]:


def A1toij(a1formula):
    r1c1absolute = xw.apps.active.api.ConvertFormula(a1formula, 1, -4150, 1)
    return tuple(int(m) for m in re.findall(r'\d+', r1c1absolute))


# In[4]:


workingdir = tkfd.askdirectory(parent = root)
if not workingdir:
    sys.exit(130)


# In[5]:


wb = xw.Book(os.path.join(workingdir, 'form.xlsx'))
template_filepath = os.path.join(workingdir, 'template.xlsx')
output_filename_template = 'page-%05d.xlsx'


# In[6]:


label = wb.sheets['Records'].range('records[#見出し]').value
records = pd.DataFrame([dict(zip(label, row)) for row in wb.sheets['Records'].range('records').value])
records


# In[7]:


configurations = dict((row[0], tuple(row[1:])) for row in wb.sheets['Configurations'].range('configurations').value)
configurations


# In[8]:


page_settings = dict((row[0], int(row[1])) for row in wb.sheets['Page Settings'].range('page_settings').value)
for k, v in page_settings.items():
    vars()[k] = v
page_settings


# In[9]:


wb.close()


# In[10]:


records_per_page = record_rows_per_page * record_columns_per_page
pages = math.ceil(len(records) / records_per_page)


# In[11]:


app = xw.App()
for i, j, k in itertools.product(range(pages), range(record_rows_per_page), range(record_columns_per_page)):
    n = records_per_page * i + record_columns_per_page * j + k
    if n >= len(records):
        break
    if j == 0 and k == 0:
        if i > 0:
            wb.save()
            wb.close()
        new_filename = os.path.join(workingdir, output_filename_template % (i + 1))
        shutil.copy(template_filepath, new_filename)
        wb = app.books.open(new_filename)
        ws = wb.sheets.active
    for col, (a1, method) in configurations.items():
        s, t = A1toij(a1)
        s += j * cells_per_row
        t += k * cells_per_column
        val = records.at[n, col]
        ws.cells(s, t).value = val
        #print(s, t, val)
app.quit()


# In[ ]:




