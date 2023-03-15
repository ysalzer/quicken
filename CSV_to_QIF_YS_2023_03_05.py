'''**************************************************************************/
#
#    @file    CSV_to_QIF.py
#    @author   Mario Avenoso (M-tech Creations)
#    @license  MIT (see license.txt)
#
#    Program to convert from a CSV to a QIF file using a definitions
#    to describe how the CSV is formatted
#
#    @section  HISTORY
#    v0.2 - Added payee ignore option  1/18/2016 feature update
#    v0.1 - First release 1/1/2016 beta release
#
#     https://github.com/M-tech-Creations/CSV-to-QIF 
#     http://mario.mtechcreations.com/programing/csv-to-qif-python-converter/
#     https://en.wikipedia.org/wiki/Quicken_Interchange_Format
#
'''#**************************************************************************/


import os
import sys
import re
import csv
import pandas as pd
import numpy as np

#%%
#SHITTY DEBUGIN

# deff_ = 'Yahav_credit_card_to_quicken.def'

# inf_ = 'Mastercard_3982_2022_12_for_Quicken.csv'

# outf_ = filelocation_ ='Mastercard_3982_2022_12_for_Quicken.qif'

# path = os.path.dirname(os.path.realpath("__file__"))
# path_parent = os.path.dirname(os.getcwd())
# path_source_files = path + '\files'

# path = 'D:\Documents\Personal\9-Quicken\Code\CSV-to-QIF-master\CSV-to-QIF-master\\'
#%%
def writeFile(date_,amount_,memo_,payee_, cat_, path_files, outf_):
    outFile = open(path_files + '\\' + outf_,"a")  #Open file to be appended
    outFile.write("!Type:Cash\n")  #Header of transaction, Currently all set to cash
    outFile.write("D")  #Date line starts with the capital D
    outFile.write(date_)
    outFile.write("\n")

    outFile.write("T")  #Transaction amount starts here
    outFile.write(amount_)
    outFile.write("\n")

    outFile.write("M")  #Memo Line
    outFile.write(memo_)
    outFile.write("\n")
    
    outFile.write("L")  #Memo Line
    outFile.write(cat_)
    outFile.write("\n")
    
    outFile.write("P")  #Payee line
    outFile.write(payee_)
    outFile.write("\n")

    outFile.write("^\n")  #The last line of each transaction starts with a Caret to mark the end
    outFile.close()
  #%%

def convert(inf_,outf_, path_files): #will need to receive input csv and def file
    
    if not os.path.isfile(path_files + '\\' + outf_):
        
        df = pd.read_csv(path_files + '\\' + inf_)
        df.pop(df.columns[0])
        df = df.replace(np.nan, '')
        
        for index, row in df.iterrows():
            # print(row['date'],row['amount'],row['memo'],row['payee'], row['category'],)
            writeFile(row['date'],str(row['amount']),str(row['memo']),row['payee'], str(row['category']), path_files, outf_)  #export each row as a qif entry
            

# https://en.wikipedia.org/wiki/Quicken_Interchange_Format
#
