import pandas as pd

# data load
df = pd.read_excel('data_sheet.xlsx', engine='xlrd')

# Select Total rows 
total = df[df['Area'] == 'Total']

print(total.shape)  
print(total.head()) 

