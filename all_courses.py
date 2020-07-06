#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 29 21:25:58 2020

@author: haopeng
"""

'''
Generage a single csv file and html page which include all information of NetAppU ILT Course Materials
The content is the same as page
https://netapp.sharepoint.com/sites/NetAppU/CPD/ILT/Lists/ReleasedProducts1/Course%20and%20Kit.aspx
'''

from shareplum import Site
from shareplum import Office365
#from shareplum.site import Version
from pandas import DataFrame
import os

#replace the domain_user and domian_password with your real ones
domain_user = 'your_userID@netapp.com'
domian_password = 'your_password'

# replace the path with your path
path = '/Users/haopeng/instructor/OneDrive - NetApp Inc/My_Project/all_courses'

authcookie = Office365('https://netapp.sharepoint.com', username=domain_user, password=domian_password).GetCookies()
site = Site('https://netapp.sharepoint.com/sites/NetAppU/CPD/ILT/', authcookie=authcookie)

# find the list name use Chrome --> inspect --> nework --> XHR --> Preview --> ListName
sp_list = site.List('77CF21E7-1FDF-4C5C-9EF4-01C2E14A442D')

data = sp_list.GetListItems('All Items')

df = DataFrame.from_dict(data)

file_name = 'all_courses.csv'
fullPath = os.path.join(path, file_name)

#render to csv
df.to_csv (fullPath, index = False, header=True)

#render dataframe as html
#escape bool, default True
#Convert the characters <, >, and & to HTML-safe sequences.
html = df.to_html(escape=False,classes='table table-striped',justify='center')

#write html to file
html_name = 'index.html'
html_fullPath = os.path.join(path, html_name)

temp_file = 'template.html'
temp_fullPath = os.path.join(path, temp_file)
temp_file = open(temp_fullPath, "r")

text_file = open(html_fullPath, "w",encoding="utf-8")
for line in temp_file:
    text_file.write(line)

text_file.write(html)
text_file.writelines('</body>')
text_file.writelines('</html>')
text_file.close()
temp_file.close()


