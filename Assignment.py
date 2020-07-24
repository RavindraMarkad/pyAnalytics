# -*- coding: utf-8 -*-
"""
Created on Tue Jul 21 11:17:18 2020

@author: Ravindra Markad
"""
import pandas as pd
import numpy as np

data = pd.read_csv('denco.csv')#import file

data.head()#checking import by loading headers
type(data)#checking data type of imported file
#%%

data.sort_values(by=['custname']) #sorting by customer name
group = data.groupby('custname').aggregate({'revenue': [np.sum,'count']})# summing repeated customers

sh1 = group.nlargest(5,[('revenue',   'count')])# top 5 customers count

#%%
sh2 = group.sort_values([('revenue',   'sum')], ascending=False)#sort by revenue descending

#%%
group_part_1 = data.groupby('partnum').aggregate({'revenue':[np.sum],'margin':[np.sum]})
group_part_1.columns
p1=group_part_1.sort_values([('revenue', 'sum')],ascending=False) # sorted by part number revenue
p1.head(5)

#%%
group_part_2 = data.groupby('partnum').aggregate({'margin':[np.sum]})
p2=group_part_2.sort_values([( 'margin', 'sum')],ascending=False)
p2.head(5) # top 5  part by margin

#%%
# writing files in excel
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
# Write each dataframe to a different worksheet. you could write different string like above if you want
group.to_excel(writer, sheet_name='group by customer')
sh1.to_excel(writer, sheet_name='top 5 cust by count')
p1.to_excel(writer, sheet_name='Revenue by part')
p2.to_excel(writer, sheet_name='Margin by part')
# Close the Pandas Excel writer and output the Excel file.
writer.save()

