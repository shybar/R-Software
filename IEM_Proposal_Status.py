import os		# Use for system Path location
os.getcwd()       # Prints the current working directory
path="D:\SHYAMA_WORKING\ZUTIL_PROPOSAL" 	# Set current working path
os.chdir(path)  # Provide the new path here
os.getcwd()       # Prints the current working directory


import pandas as pd					# Use for Pandas method
from pandas import ExcelWriter		# Use for Read Excel
from pandas import ExcelFile		# Use for Excel

# Read Excel file
jc=pd.read_excel('JC_NAME.xlsx',sheet_name='Sheet1')

# Read Excel file
file1=pd.read_excel('ZUTIL_PROPOSAL_P92.xlsx',sheet_name='Sheet1')
df1=file1[['Proposal Doc. No.','Proposal Date','Site ID','Consumer number', 'Meter Number','Vendor Code','Vendor Name','Vendor Invoice No.','Amount','Scroll No.','JC Site Id']]

# Read Excel file
file2=pd.read_excel('ZUTIL_MASTER_P92.xlsx',sheet_name='Sheet1')
df2= file2[['Consumer number','Meter No','Tax Code','Site ID']]

# Read Excel file
file3=pd.read_excel('YEXP_PROPOSAL_P92.xlsx',sheet_name='Sheet1')
df3= file3[['Proposal Number','Short Text','Short Text.1']].rename(columns={'Short Text':'Proposal Status'}).rename(columns={'Short Text.1':'Short Text'})

# Create marge with Left Joing into two files, Drop column, Rename column
m1=pd.merge(left=df1,right=df2, how='left', left_on=['Consumer number', 'Meter Number','Site ID'], right_on=['Consumer number','Meter No', 'Site ID']).drop(columns=['Meter No']).rename(columns={'JC Site Id':'JC SAP ID'})

# Create marge with Left Joing into two files like Vlookup
m2=pd.merge(left=m1,right=df3, how='left', left_on=['Proposal Doc. No.'], right_on=['Proposal Number'])

# Create marge with Left Joing into two files, Drop column
m3= pd.merge(left=m2,right=jc, how='left', on=['JC SAP ID']). drop(columns=['Proposal Doc. No.'])

# Arranging columns
P92_Final = m3[['Consumer number','Meter Number','Vendor Code','Vendor Name','Vendor Invoice No.','Amount','Site ID','JC SAP ID','JIO CENTRE NAME','Proposal Date','Proposal Number','Proposal Status','Scroll No.','Tax Code','Short Text']]

# Convert the dictionary into DataFrame
df=pd.DataFrame(P92_Final)

import time			# Use for Pandas method
datetime = time.strftime("%Y%m%d_%H%M")		# Use for current datetime format
FileName='P92_IEM_Proposal_'+ datetime +'.xlsx' 	# Use export filename with current dataframe

# Create a Pandas Excel writer using XlsxWriter as the engine.
#writer = pd.ExcelWriter('Proposal_IEM-P92.xlsx', engine='xlsxwriter')
writer = pd.ExcelWriter(FileName, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1',index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

#--------Use for P91 Proposal

# Read Excel file
file4=pd.read_excel('ZUTIL_PROPOSAL_P91.xlsx',sheet_name='Sheet1')
df4=file4[['Proposal Doc. No.','Proposal Date','Site ID','Consumer number', 'Meter Number','Vendor Code','Vendor Name','Vendor Invoice No.','Amount','Scroll No.','JC Site Id']]

# Read Excel file
file5=pd.read_excel('ZUTIL_MASTER_P91.xlsx',sheet_name='Sheet1')
df5= file5[['Consumer number','Meter No','Tax Code','Site ID']]

# Read Excel file
file6=pd.read_excel('YEXP_PROPOSAL_P91.xlsx',sheet_name='Sheet1')
df6= file6[['Proposal Number','Short Text','Short Text.1']].rename(columns={'Short Text':'Proposal Status'}).rename(columns={'Short Text.1':'Short Text'})

# Create marge with Left Joing into two files, Drop column, Rename column
m4=pd.merge(left=df4,right=df5, how='left', left_on=['Consumer number', 'Meter Number','Site ID'], right_on=['Consumer number','Meter No', 'Site ID']).drop(columns=['Meter No']).rename(columns={'JC Site Id':'JC SAP ID'})

# Create marge with Left Joing into two files like Vlookup
m5= pd.merge(left=m4,right=df6, how='left', left_on=['Proposal Doc. No.'], right_on=['Proposal Number'])

# Create marge with Left Joing into two files, Drop column
m6= pd.merge(left=m5,right=jc, how='left', on=['JC SAP ID']). drop(columns=['Proposal Doc. No.'])

# Arranging columns
P91_Final = m6[['Consumer number','Meter Number','Vendor Code','Vendor Name','Vendor Invoice No.','Amount','Site ID','JC SAP ID','JIO CENTRE NAME','Proposal Date','Proposal Number','Proposal Status','Scroll No.','Tax Code','Short Text']]

# Convert the dictionary into DataFrame
df=pd.DataFrame(P91_Final)

FileName='P91_IEM_Proposal_'+ datetime +'.xlsx' 	# Use export filename with current dataframe

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(FileName, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1',index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
