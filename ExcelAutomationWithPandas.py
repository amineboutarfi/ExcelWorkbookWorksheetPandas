# import pandas
import pandas as pd


################### CSV to EXCEL

# read csv file
read_file = pd.read_csv('bestsellers with categories.csv')
# save dataframe into excel
read_file.to_excel('bestsellersWithCaregories.xlsx', index = None, header=True)


################### ONE WORKSHEET TO MULTIPLE WORKBOOKS

# read excel workbook - if we have one worksheet
df_file = pd.read_excel('bestsellersWithCaregories.xlsx')

# read excel workbook - if multiple worksheets
df_file = pd.read_excel('bestsellersWithCaregories.xlsx', sheet_name='sheetname')

# exctract 2015 data
results_2015 = df_file[df_file['Year'].apply(str).str.match('2015')]
# exctract 2016 data
results_2016 = df_file[df_file['Year'].apply(str).str.match('2016')]
# exctract 2017 data
results_2017 = df_file[df_file['Year'].apply(str).str.match('2017')]

# save 2015 data into excel workbook
results_2015.to_excel('bestsellersWithCaregories2015.xlsx', index = None, header=True)
# save 2016 data into excel workbook
results_2016.to_excel('bestsellersWithCaregories2016.xlsx', index = None, header=True)
# save 2017 data into excel workbook
results_2017.to_excel('bestsellersWithCaregories2017.xlsx', index = None, header=True)

################### ONE WORKSHEET TO MULTIPLE WORKSHEETS

df_file = pd.read_excel('bestsellersWithCaregories.xlsx')

results_2015 = df_file[df_file['Year'].apply(str).str.match('2015')]
results_2016 = df_file[df_file['Year'].apply(str).str.match('2016')]
results_2017 = df_file[df_file['Year'].apply(str).str.match('2017')]


writer = pd.ExcelWriter('final.xlsx')

results_2015.to_excel(writer,'2015', index = None)
results_2016.to_excel(writer,'2016', index = None)
results_2017.to_excel(writer,'2017', index = None)

writer.save()

################### MULTIPLE WORKSHEETS TO ONE WORKBOOK

df_file_2015 = pd.read_excel('bestsellersWithCaregories2015.xlsx')
df_file_2016 = pd.read_excel('bestsellersWithCaregories2016.xlsx')
df_file_2017 = pd.read_excel('bestsellersWithCaregories2017.xlsx')

df_2015_2017 = pd.concat([df_file_2015, df_file_2016, df_file_2017])

df_2015_2017.to_excel('bestsellersWithCaregories201520162017.xlsx', index = None, header=True)
