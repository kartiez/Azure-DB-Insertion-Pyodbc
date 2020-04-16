import pyodbc
import os
from IPython.display import clear_output
server = 'YOUR SERVER NAME' 
database = 'YOUR DB NAME' 
username = 'USERNAME' 
password = 'PASSWORD@123' 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
from ipywidgets import IntProgress
from IPython.display import display
import time

import pandas as pd


# In[11]:


df = pd.read_excel('FILE.XLSX')
df.head()


# In[55]:


max_count = 300


# In[65]:


tquery = "TRUNCATE TABLE <TABLNAME>"
cursor.execute(tquery)
print("Database Truncated!")


# In[66]:


f = IntProgress(min=0, max=max_count) 
display(f)
for index, row in df.iterrows():
    row['Review'] = row['Review'].replace("'","")
    query = "INSERT INTO <TABLENAME> VALUES ({},'{}',{},{},'{}','{}','{}','{}','{}',{},{},'{}','{}')".format(row['sno'],row['Location'],row['Latitude'],row['Longitude'],row['Postal_Code'],row['Current_Status'],row['Contact_No'],row['Open'],row['Close'],row['Number_of_Reviews'],row['Overall_Rating'],row['Review'],row['polarity'])
    cursor.execute(query)
    cnxn.commit()
    f.value += 1 
print("Insert Operation Completed")


# In[ ]:




