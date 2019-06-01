# Importing required modules 
import numpy as np		# Use for Numpy module for number	
import pandas as pd		# Use for Pandas module
import zipfile		# Use for ZIP file
import os		# Use for system Path location

path="D:\\SHYAMA_WORKING\\NOC_REPORT\\Python_Report" 	# Current working path
os.chdir(path)  # Set Current working path
os.getcwd()       # Prints the current working directory

# Read zip file
zip= zipfile.ZipFile('Access availability Report_30052019.zip')

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

# ExportFileName
ExportFileName='NOC_Report_'+fname[27:-5]+'.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(ExportFileName, engine='xlsxwriter')

# Write each dataframe to a different worksheet.
enb.to_excel(writer, sheet_name='EnB',index=False)
isc.to_excel(writer, sheet_name='ISC',index=False)
osc.to_excel(writer, sheet_name='OSC',index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Close zip file
zip.close()






	  