# Importing required modules 
import numpy as np		# Use for Numpy module for number	
import pandas as pd		# Use for Pandas module
import zipfile		# Use for ZIP file
import os		# Use for system Path location

path="D:\\SHYAMA_WORKING\\NOC_REPORT\\Python_Report" 	# Current working path
os.chdir(path)  # Set Current working path
os.getcwd()       # Prints the current working directory

# Read zip file
zip= zipfile.ZipFile('Access availability Report_27062019.zip')

# Read file from zip
fname=zip.namelist()
fname=str(fname)

# initializing bad_chars_list
bad_chars = ['"', '[', ']', "'"]

# using replace() to  
# remove bad_chars 
for i in bad_chars : 
    fname = fname.replace(i, '')

# open file from zip 
fopen=zip.open(fname)

enb=pd.read_excel(fopen,sheet_name='PAN_INDIA')
isc=pd.read_excel(fopen,sheet_name='Indoor Small cell')
osc=pd.read_excel(fopen,sheet_name='Outdoor Small cell')
enb=enb.loc[enb['Circle'] == 'Kolkata']
isc=isc.loc[isc['R4G'] == 'Kolkata']
osc=osc.loc[osc['R4G'] == 'Kolkata']

# drop a column based on column name
enb=enb.drop(["Sr NO"], axis = 1)

# Insert a column at a specific column index
idx = 0   #  Index value 0 meance 1st column

# Inser new column 'Date' in 1st column with value= ZIP file date
enb.insert(loc=idx, column='Date', value=fname[27:-5])  ## substrick from fname
isc.insert(loc=idx, column='Date', value=fname[27:-5])  ## substrick from fname
osc.insert(loc=idx, column='Date', value=fname[27:-5])  ## substrick from fname


# ExportFileName
Export_EnB='NOC_Report_EnB_'+fname[27:-5]+'.csv'
Export_ISC='NOC_Report_ISC_'+fname[27:-5]+'.csv'
Export_OSC='NOC_Report_OSC_'+fname[27:-5]+'.csv'

# Create a Pandas Excel writer using XlsxWriter as the engine.
#writer = pd.ExcelWriter(ExportFileName, engine='xlsxwriter')

# Write each dataframe to a different CSV file.
enb.to_csv(Export_EnB,index=False)
isc.to_csv(Export_ISC,index=False)
osc.to_csv(Export_OSC,index=False)

# Close the Pandas Excel writer and output the Excel file.
#writer.save()

# Close zip file
zip.close()






	  