# -*- coding: utf-8 -*-
"""
Created on Fri Mar 10 17:24:26 2023

@author: salzer
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

from datetime import datetime


#%%

path = os.path.dirname(os.path.realpath(__file__))
path_parent = os.path.dirname(os.getcwd())

path_files = path + '\\files\\'
path_dictionary_files = path + '\\files_dictionary\\'


#%%

xls_list = glob.glob(path_files + 'Mastercard_*.xls')

appended_df = pd.DataFrame()
unique_df = pd.DataFrame()

for xls_file in xls_list:
    print(xls_file)
   # NIS transactions

    df = pd.read_excel(xls_file, sheet_name = 'Transactions')

    # Drop first 6 rows of dataframe
    df = df.drop(np.linspace(0,3,4).astype(int)).reset_index()
    df.drop(columns=df.columns[0], axis=1,  inplace=True)
     
      #Replace headers with first row, drop first row
    df = df.rename(columns=df.iloc[0]).drop(df.index[0]) 

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

    try:
           #Remove foreign currency transactions
        USD_transtactions_index = df[df['DateTime'] == 'עסקאות בחו˝ל'].index.values[0]
        df = df[:USD_transtactions_index-2]
        print('Foreign currency transactions')
    except:
        #Remove two last rows
        df = df[:-2]
        print('No foreign currency transactions')

    appended_df = appended_df.append(df)

unique_df['hebrew'] = appended_df['payee'].unique()
# uniqe_df['category'] = ''
# uniqe_df['payee'] = ''

try:
    filled_unique_df = pd.read_excel(path_dictionary_files + '/dictionary_of_category_payee.xlsx')
    filled_unique_df = filled_unique_df[['hebrew','payee','category']]
except:
    print('No previouse lists')

united_unique_df = pd.merge(filled_unique_df, unique_df, on = 'hebrew', how = 'outer')

timestamp = datetime.now().strftime("%Y_%m_%d-%I_%M")

united_unique_df.to_excel(path_dictionary_files + '/dictionary_of_category_payee_%s.xlsx' % timestamp)
