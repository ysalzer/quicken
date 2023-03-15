# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 18:09:36 2023

@author: salzer

# Original file named convert_credit_csv_to_qif.py
"""


import datetime as dt
import glob
import os
import tkinter as tk
from tkinter import Tk  # from tkinter import Tk for Python 3.x

import numpy as np
import pandas as pd
import pylab as pyl
import scipy as sy
import seaborn as sns
from matplotlib import pyplot as plt
# denoising_data_with_FFT
from pandastable import Table
from scipy.stats import stats

import re
import csv
#import quiffen

import sys
sys.path.append(r"D:\Documents\Personal\9-Quicken\Code")

from CSV_to_QIF_YS_2023_03_05 import writeFile, convert

#%%

path = os.path.dirname(os.path.realpath(__file__))
path_parent = os.path.dirname(os.getcwd())

path_files = path + '\\files\\'
path_dictionary_files = path + '\\files_dictionary\\'

#%% 
###                    MasterCard credit card files 
###            Convert xls credit card files to csv, and then to qif
### ------------------------------------------------------------------------

    
        
def Mastercard_to_qif(xls_file):
        # NIS transactions
        df = pd.read_excel(xls_file, sheet_name = 'Transactions')
        
        # Drop first 6 rows of dataframe
        df = df.drop(np.linspace(0,3,4).astype(int)).reset_index()
        df.drop(columns=df.columns[0], axis=1,  inplace=True)
    
        # Replace headers with first row, drop first row
        df = df.rename(columns=df.iloc[0]).drop(df.index[0])
       
        #Extract information of file [card number, month, year]
        #REF: https://stackoverflow.com/questions/45234640/how-to-extract-part-of-a-filename-used-in-a-python-script-as-an-argument
        file_name = xls_file.split('\\')[-1].split('.')[0]
        re_input = 'Mastercard_(?P<card>\d+)_Export_(?P<month>\d+)_(?P<year>\d+)'
        match = re.match(re_input, file_name)
        year = match.group('year')
        month = match.group('month')
        card = match.group('card')
    
        #Rename required hebrew column headers
        columns_hebrew_reduced = [
    'תאריך רכישה', 
'שם בית עסק', 
'מספר שובר', 
'סכום חיוב',
       'פירוט נוסף',
'סכום עסקה', 
                  ]
    
        df = df[columns_hebrew_reduced]
    
        columns_english_reduced = ['DateTime',
                                   'payee',
                          'transaction id',
                                'amount',
                                'memo',
               'memo full amount',
                      ]    
        
        columns_dictionary = dict(zip(columns_hebrew_reduced, columns_english_reduced))
        df.rename(columns = columns_dictionary, inplace = True)
       
        # If foreign transactions reported
        # remove transactions from NIS dataframe
        # save transactions to new USD dataframe

        try:
            USD_transtactions_index = df[df['DateTime'] == 'עסקאות בחו˝ל'].index.values[0]
            df_USD = df[USD_transtactions_index+1:-1]
            df = df[:USD_transtactions_index-2]
            df_USD.rename(columns = columns_dictionary, inplace = True)
            df_USD['date'] = pd.to_datetime(df_USD['DateTime']).dt.strftime('%m/%d/%Y')

        except:
                  #Remove two last rows
            df = df[:-2]
        
        df['date'] = pd.to_datetime(df['DateTime']).dt.strftime('%m/%d/%Y')

       
        # Read hebrew:payee:category csv to dataframe
        df_cat_payee_read = pd.read_excel(path_dictionary_files + '/dictionary_of_category_payee.xlsx')
        df_cat_payee_read = df_cat_payee_read[['hebrew', 'payee', 'category']]

        # Generate  hebrew:payee dictionary
        payee_dictionary = dict(zip(df_cat_payee_read['hebrew'],df_cat_payee_read['payee'] ))
        
        # Generate  payee:category dictionary
        category_dictionary = dict(zip(df_cat_payee_read['payee'],df_cat_payee_read['category']))
        
        # Replace Hebrew with English payee
        df['payee'].replace(payee_dictionary, inplace = True)
        # Add Category
        df['category'] = df['payee'].map(category_dictionary)
        

        # Merge 'memo' and 'memo full amount' if 'memo' is not empty, suggesting multiple payments        
        df['memo full amount'] = df['memo full amount'].astype(str)       
        df['memo'] = df.apply(lambda row: np.nan  if pd.isnull(row['memo']) else row["memo full amount"]+ ":" + row["memo"], axis = 1)
        df['memo'] = df.apply(lambda row: np.nan  if pd.isnull(row['memo']) else str(row["memo full amount"]), axis = 1)
        
        # Remove un-needed columns
        df = df[['date', 'payee',  'category', 'transaction id', 'amount', 'memo']]
        
        #Amount negative (payed, not received)
        df.amount = df.amount*-1
    
        
        # Generate file neames
        xls_file = '\Mastercard_{}_{}_{}_for_Quicken.xlsx'.format(card, year, month)
        csv_file = '\Mastercard_{}_{}_{}_for_Quicken.csv'.format(card, year, month)
        qif_file = '\Mastercard_{}_{}_{}_for_Quicken.qif'.format(card, year, month)
        usd_qif_file = '\Mastercard_USD_{}_{}_{}_for_Quicken.qif'.format(card, year, month)
        usd_csv_file = '\Mastercard_USD_{}_{}_{}_for_Quicken.csv'.format(card, year, month)

        #Save to CSV files
        df.to_csv(path_files + csv_file)
        
        #Covert CSV file to QIF file
        convert(csv_file, qif_file, path_files)


        # If USD transactions exist
        # Save to CSV
        # Convert to QIF
        try:
            df_USD.to_csv(path_files + usd_csv_file)
            convert(usd_csv_file, usd_qif_file, path_files)
        except:
            print("No USD transactions")
       


#%%

xls_list = glob.glob(path_files + 'Mastercard_*.xls')

for xls_file in xls_list:
    print(xls_file)
    Mastercard_to_qif(xls_file)
    
    
    
