#\*\*\*\*Create a new month sheet\*\*\*\* 
#Purpose of program: 
#adding a new sheet containing a list of companies that worked with in the following month
#import filled tax amount of companies from a downloaded excel sheet from taxation infrastructure (hometax.go.kr)

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

#1.enter specific month  
month = input("enter month to add (ex. 2020.09) : ")
#2.excel file name, downloaded from hometax.go.kr 
hometax_file = input("enter hometax file name (ex. hometax_july_dec.xls) : ")
#3.transaction history file name  
status_file = input("enter status file name (ex. status.xlsx) : ")
#4.select date range
start_date = str(month.replace('.', '-')) + "-01"
end_date = end_of_month(month)

def end_of_month(m): 
  m = int(month.split('.')[1])
  if m == 2: 
    end_day = "-28"
  elif m <= 7 and m % 2 == 1: 
    end_day = "-31"
  elif m <= 7 and m % 2 == 0: 
    end_day = "-30"
  elif m > 7 and m % 2 == 0: 
    end_day = "-31"
  elif m > 7 and m % 2 == 1: 
    end_day = "-30"
  end_date = str(month.replace('.','-')) + end_day
  return end_date

#import a list of partners worked with in the following month 
status_f = pd.read_excel(status_file, sheet_name = month, skiprows = 2)
status_df = status_f['company_name']
company_list = status_df.tolist()
month_company_list = sorted(list(set([x.strip(' ') for x in company_list if x == x]))) #remove space in element and x == x for removing nan in the list 

#creation of dictionary {company_name: company_deposit_name}  
company_depositname_file = open("company_deposit_name.txt")
file_lines = company_depositname_file.readlines()
company_depositname_dict = {}
for line in file_lines: 
  key_value = line.split()
  company_depositname_dict[key_value[0]] = key_value[1]

#import filled tax history from hometax excel file
hometax_f = pd.read_excel(hometax_file, skiprows = 5)
hometax_df = hometax_f[['Date','company','tax_total']]
#display selected range 
hometax_df['Date'] = pd.to_datetime(hometax_df['Date'])  
mask = (hometax_df['Date'] >= start_date) & (hometax_df['Date'] <= end_date) 
hometax_df = hometax_df.loc[mask]

#creation of DataFrame containing list of companies and its tax amount
month_company_df = pd.DataFrame(index = [1,2,3,4,5,6,7,8,9,10], columns=month_company_list)
for company in month_company_list:
  try: 
    deposit_name = company_depositname_dict[company]
    cost_list = (hometax_df[hometax_df['company'].str.contains(deposit_name, na=False)]['tax_total']).tolist()
    if len(cost_list) < 10: 
      cost_list += [0]*(10-len(cost_list))
    month_company_df[company] = cost_list
  except: 
    None
#note: 
#company with 0 filled column - not registered in hometax 
#company with NAN column - invaild company_deposit_name provided 

#add final dataFrame as a sheet to existing payment history excel file
from openpyxl import load_workbook
book = load_workbook('payment_history.xlsx')
writer = pd.ExcelWriter('payment_history.xlsx', engine = 'openpyxl')
writer.book = book
month_company_df.T.to_excel(writer, sheet_name = month)
writer.save()
writer.close()

print(month + " sheet created: download 'payment history.xls'")