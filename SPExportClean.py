'''
Created on 7 Jun 2015

@author: Mike
'''

import csv     
import sys  
import os
import re



try:
    in_file_name = sys.argv[1]    
    out_file_name = sys.argv[2]
except:
    print("Usage: python sharepoint_export_clean.py <input_csv_file> <output_xlsx_file>")
    print("Where <input_csv_file is export from sharepoint")
    print("      <output_csv_file> will export with usernames cleaned up (hash and digit strings removed)")
    sys.exit(0)


# Load the input file
try:
    print("Opening File...")
    input_file = open(in_file_name, encoding='utf-16') # opens the csv file
except IOError:
    print ('Error. Cannot open', in_file_name)
    sys.exit(0)

print("Reading File...")
reader = csv.reader(input_file,delimiter='\t')  # creates the reader 
raw_data = [r for r in reader]

# Create output file
try:
    print("Opening Output File...")
    out_file = open(out_file_name, encoding='utf-16',mode='w') # opens the csv file
except IOError:
    print ('Error. Cannot open', out_file_name)
    sys.exit(0)

# Create Reg Exp
regexp=re.compile(';#[0-9]+')
regexp2=re.compile(';#')
    
for r in raw_data:
    print(r)
    
    for i in r:
        out_file.write(regexp2.sub(';',regexp.sub('',i))+'\t')
    
    out_file.write('\n') # End of record
    

input_file.close()      # close input file    
out_file.close() # close output file            
print("Done")
