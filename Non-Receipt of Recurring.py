
# Importing required modules 
import os		# Use for system Path location
import numpy as np		# Use for Numpy module for number	
import pandas as pd		# Use for Pandas module
from pandas import ExcelFile		# Use for Excel
from pandas import ExcelWriter		# Use for Read Excel
import time			# Use for Pandas method

path="D:\\SHYAMA_WORKING\\RECURRING_BILL_NOT_RECEIVED\\Python_Report" 	# Current working path
os.chdir(path)  # Set Current working path
os.getcwd()       # Prints the current working directory

# Create dataframe
d={'JC_SAPID' : ['I-KO-KLKT-JCO-0001','I-KO-KLKT-JCO-0002','I-KO-KLKT-JCO-0003','I-KO-KLKT-JCO-0004','I-KO-KLKT-JCO-0005','I-KO-KLKT-JCO-0006','I-KO-KLKT-JCO-0007','I-KO-KLKT-JCO-0008','I-KO-KLKT-JCO-0009','I-KO-KLKT-JCO-0010','I-KO-KLKT-JCO-0011','I-KO-KLKT-JCO-0012','I-KO-KLKT-JCO-0013','I-KO-KLKT-JCO-0014'],
     'JC_NAME':['CHANDAN NAGAR','SERAMPORE','BARRACKPORE','DUNLOP','BARASAT','RAJARHAT','HOWRAH','CENTRAL AVENUE','ALIPORE','PARKSTREET','KASBA','JAMES LONG SARANI','KAMALGAZI','NAIHATI'],
	 'CMP_LEAD':['Bikash Pradhan','Bikash Pradhan','Ramakanta Mohapatra','Ramakanta Mohapatra','Ramakanta Mohapatra','Ramakanta Mohapatra','Bikash Pradhan','Ramakanta Mohapatra','Raju Dubey','Raju Dubey','Ramakanta Mohapatra','Raju Dubey','Ramakanta Mohapatra','Ramakanta Mohapatra']}

# Convert to pandas dataframe
df1=pd.DataFrame(d)

# Read Excel file
f1=pd.read_excel('Non-Receipt of Recurring Bill P91.xlsx',0)
f2=pd.read_excel('Non-Receipt of Recurring Bill P92.xlsx',0)

# Concatenating the dataframes
f3=pd.concat([f1,f2])

# Create marge with Left Joing into two files, Drop column
m1=pd.merge(left=f3,right=df1, how='left', left_on=['JC Site ID'], right_on=['JC_SAPID']).drop(columns=['JC_SAPID'])


# Multiple if else conditions in pandas dataframe and derive multiple columns based on column		   
def Ageing(x):
    if (x['Over Due Days'] < 1):
        return 'Nil'
    elif (x['Over Due Days'] > 0 and (x['Over Due Days'] < 8)):
        return '1-7 Days'
    elif (x['Over Due Days'] > 7 and (x['Over Due Days']<= 15)):
        return '8-15 Days'
    elif (x['Over Due Days'] > 15):
        return '>15 Days'
	
Pending = m1.assign(Ageing_Slab=m1.apply(Ageing, axis=1))

# Pivot Table
Summary=pd.pivot_table(Pending,
           index=['JC_NAME','CMP_LEAD'],
		   columns='Ageing_Slab',
           margins=True,
           aggfunc=len,   # len to get count, # aggfunc=np.sum for sum calculation
           values='Site ID')

# Convert to pandas dataframe
Summary=pd.DataFrame(Summary)

# Use for current datetime format
datetime = time.strftime("%Y%m%d_%H%M")		
		   
# ExportFileName
ExportFileName='Non-Receipt_Recurring_Bill_'+ datetime +'.xlsx'	

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(ExportFileName, engine='xlsxwriter')

# Write each dataframe to a different worksheet.
Summary.to_excel(writer, sheet_name='Summary_sheet')
Pending.to_excel(writer, sheet_name='Details',index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
	   