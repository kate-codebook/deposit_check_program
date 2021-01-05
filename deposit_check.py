#look up status of payment 
from pandas import DataFrame, read_csv
import matplotlib.pyplot as plt
import pandas as pd 
import numpy as np
import csv
import string
import warnings
from pandas.core.common import SettingWithCopyWarning
warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)
warnings.filterwarnings("ignore", 'This pattern has match groups')
import xlrd
import sys

#7 files required to run the program 
#1.4 bank statements from accounts 
nhmain_file = r'nh_main.xls'    
nhbranch_file = r'nh_branch.xls'
j_file = r'j_account.xls' 
ibkb_file = r'ibkb_file'    
#2.home tax excel file
hometax_file = r'hometax_july_dec.xls' 
#3. payment history excel file
loc = ("payment_history.xlsx")
#4. company_deposit_name.txt
#5. select period range
start_date = '2020-07-01'
end_date = '2020-12-31'

#nh bank main statement
nhmain_f = pd.read_excel(nhmain_file, skiprows = 9)
nhmain_df = nhmain_f[['date', 'deposit$', 'depositor']]
#nh bank branch statement
nhbranch_f = pd.read_excel(nhbranch_file, skiprows = 9)
nhbranch_df = nhbranch_f[['date', 'deposit$', 'depositor']]
#ibk_indivisual account statement    
j_f = pd.read_excel(j_file)
j_df = j_f[['date', 'deposit$', 'depositor']]
#ibk_corp account statement    
ibkb_f = pd.read_excel(ibkb_file)
ibkb_df = ibkb_f[['date', 'deposit$', 'depositor']]

#import companies and depositor_names from company_deposit_name.txt and create dictionary 
company_depositname_file = open("company_deposit_name.txt")
file_lines = company_depositname_file.readlines()
company_depositname_dict = {}
for line in file_lines: 
  key_value = line.split()
  company_depositname_dict[key_value[0]] = key_value[1]

#import tax history from hometax.go.kr excel file
hometax_f = pd.read_excel(hometax_file, skiprows = 5)
hometax_df = hometax_f[['date','company','total_tax_amount']]
#display selected range 
hometax_df['date'] = pd.to_datetime(hometax_df['date'])  
mask = (hometax_df['date'] >= start_date) & (hometax_df['date'] <= end_date) #************#************
hometax_df = hometax_df.loc[mask]

#enter company name to look up the payment history
look_up_comp = 'sun-life'

#uptodate recorded-payment history 
try: 
  res = str([key for key, val in company_depositname_dict.items() if look_up_comp in val][0])
except: 
  res = look_up_comp
#find value and display rows from all the sheets in excel file
#loc = ("payment_history.xlsx")  # excel file name
wb = xlrd.open_workbook(loc)
for sheet in wb.sheet_names(): # each sheet from all sheets list 
  payment_history_df = pd.read_excel(loc, sheet_name = sheet)
  final_result_df = payment_history_df.loc[payment_history_df['Unnamed: 0'].str.contains(res, na=False)]
  print("<"+sheet+">")
  final_result_df.to_csv(sys.stdout,header=False,index=False,float_format='%d') #to remove dataframe row & col_names

#bank statements payment history  
display(nhmain_df[nhmain_df['depositor'].str.contains(look_up_comp, na=False)])
display(nhbranch_df[nhbranch_df['depositor'].str.contains(look_up_comp, na=False)])
display(j_df[j_df['depositor'].str.contains(look_up_comp, na=False)])
display(ibkb_df[ibkb_df['depositor'].str.contains(look_up_comp, na=False)])

#amount to receive from the following company 
hometax_df[hometax_df['company'].str.contains(look_up_comp, na=False)]