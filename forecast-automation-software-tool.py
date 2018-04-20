# -*- coding: utf-8 -*-
"""
Created on Mon Feb  5 22:40:04 2018

@author: Deepee
"""

#%% LIBRARY IMPORTS

import pandas as pd
import numpy as np
import math
import sys
import os
import datetime
import time
import webbrowser
from pandas import ExcelWriter

#%% FETCHING THE DIRECTORY OF THE PROGRAM AND RELEVANT FILES 

print('----------------------------------------------------------------------------------------------------------')
print('\nWelcome to FASTER!\n')

programfiles_dir = os.getcwd() + '\\'

time.sleep(2)

#%% FETCHING FILES RELEVANT TO THE PROGRAM

print('----------------------------------------------------------------------------------------------------------')

print('\nSearching for categories file....\n')
time.sleep(2)    

try:
        
    categories = pd.read_excel(programfiles_dir + 'categories.xlsx', dtype = str)
    
except:
    
    print('\nNot found. Please restart program after putting the file in appropriate location.\n')
    os._exit(0)

print('\nCategories file found and imported!\n')
time.sleep(2)

print('\nSearching for OpCos file....\n')
time.sleep(2)
    
try:

    opcos = pd.read_excel(programfiles_dir + 'opcos.xlsx', dtype = str)
    
except:
    
    print('\nNot found. Please restart program after putting the file in appropriate location.')
    os._exit(0)    

print('\nOpCos file found and imported!\n')
time.sleep(2)


#%% SETTING THE CURRENT FOLDER

print('----------------------------------------------------------------------------------------------------------')

print('\nPlease select the current folder. Available files and folders are:\n')
time.sleep(2)

ff = os.listdir(programfiles_dir)

for i in range(0,len(ff)):
    
    print(i,ff[i])
    

while True:
    
    try:
        folder_index = int(input('\nEnter a number:\n'))
        userfiles_dir = programfiles_dir + ff[folder_index] + '\\'
        userfiles = os.listdir(userfiles_dir)
        
    except:
        print('Invalid, please try again.')
        continue
    
    else:
        break
    
print('\nCurrent folder is now set. Make sure your raw forecast file is in there!')

time.sleep(2)

#%% READ IN THE FORECAST FILE

print('----------------------------------------------------------------------------------------------------------')

print('\nPlease select the csv forecast file. Available files in current folder are:\n')
time.sleep(2)

for i in range(0,len(userfiles)):
    print(i,userfiles[i])

while True:
    
    try:
        csv_index = int(input('\nEnter a number:\n'))
        
        forecast_filename = userfiles[csv_index]

        print('\nSearching for forecast file....\n')
        df = pd.read_csv(userfiles_dir + forecast_filename, dtype={'purchase_comment': str, 'part_number':str})
        print('\nForecast file found and imported! \n')
        
    except:
        
        print('\nInvalid, please try again.\n')
        continue
    
    else:
        break
            

time.sleep(2)
print('\nNow processing..\n')

#%% PREPROCESSING

print('----------------------------------------------------------------------------------------------------------')

df['receipt_date'] = pd.to_datetime(df['receipt_date'])
df['expiration_date'] = pd.to_datetime(df['expiration_date'])

# Making a dictionary of parts mapped to their categories

category_dict = {}

for col in categories.columns:
    
    for part in categories[col]:
        
        if part != 'nan':
            
            category_dict[part] = col

# Making a dictionary of categories mapped to their opcos
        
opcos_dict = {}

for col in opcos.columns:
    
    for category in opcos[col]:
        
        if category != 'nan':
            
            opcos_dict[category] = col

# Assigning categories to rows in the forecast file

categories_list = []
            
for i in range(0,len(df)):
    
    try:
        categories_list.append(category_dict[df.loc[i]['part_number']])
        
    except:
        categories_list.append('')

df['part_category'] = categories_list        

# Assigning OpCos to rows in the forecast file

opcos_list = []

for i in range(0,len(df)):
    
    try:
        opcos_list.append(opcos_dict[df.loc[i]['part_category']])
        
    except:
        opcos_list.append('')

df['opco'] = opcos_list


#%% DEFINING FUNCTIONS

# Function for calculation of outputs based on a particular forecast

def calc(weeks):
    
    dfsub = df.loc[df['usage_weeks']==weeks]
    dfsub = dfsub.loc[dfsub['expiration_date']<year_end]
    dfsub = dfsub.loc[dfsub['expiration_date']>year_start]
    dfsub_quantity = dfsub.groupby(['opco','part_category','part_number','description'])['expected_quantity'].sum()
    dfsub_unitcost = dfsub.groupby(['opco','part_category','part_number','description'])['unit_cost'].mean()
    dfsub_combined = pd.concat([dfsub_unitcost,dfsub_quantity], axis=1)
    dfsub_combined['total_cost'] = dfsub_combined['expected_quantity'] * dfsub_combined['unit_cost']
    
    exact_quantity = list(dfsub_combined['expected_quantity'])
    rounded_quantity = [int(math.ceil(x/100))*100 for x in exact_quantity]
    dfsub_combined['rounded_quantity'] = rounded_quantity
    
    exact_unitcost = list(dfsub_combined['unit_cost'])
    rounded_unitcost = [round(x,2) for x in exact_unitcost]
    dfsub_combined['rounded_unitcost'] = rounded_unitcost
    dfsub_combined['rounded_cost'] = dfsub_combined['rounded_quantity'] * dfsub_combined['rounded_unitcost']
    return dfsub_combined
    

# Function to calculate mixed forecasts

def calc_mixed():
    
    print('\nPlease select the quantity comparisons file. Available Files in Current Folder are:\n')

    userfiles = os.listdir(userfiles_dir)

    for i in range(0,len(userfiles)):
        
        print(i,userfiles[i])
        
    mixed_index = int(input('\nEnter a number:\n'))
    
    mixed_filename = userfiles[mixed_index]
    
    print('\nSearching for quantity comparisons file....\n')
    df2 = pd.read_excel(userfiles_dir + mixed_filename, index_col=[0,1,2,3], dtype={'part_number':str})
    print('\nQuantity comparisons file found and imported! \n')
    time.sleep(2)
    
    
    exact_quantity = list(df2['Chosen Qty'])
    rounded_quantity = []
    
    for x in exact_quantity:
        
        try:
            rounded_quantity.append(int(math.ceil(x/100))*100)
            
        except:
            rounded_quantity.append(0)
    
    df2['Rounded Qty'] = rounded_quantity
    
    exact_unitcost = list(df2['unit_cost'])
    rounded_unitcost = [round(x,2) for x in exact_unitcost]
    df2['Rounded Unit Cost'] = rounded_unitcost
    df2['Total Cost'] = df2['unit_cost']*df2['Chosen Qty']
    df2['Rounded Total Cost'] = df2['Rounded Qty'] * df2['Rounded Unit Cost']
    
    writer = ExcelWriter(userfiles_dir + 'Cost Calculations' +'.xlsx', engine='xlsxwriter')
    df2.to_excel(writer,'Calculations')
    
    book1 = writer.book
    format_cost_decimals = book1.add_format({'num_format':'$ #,##0.00', 'align':'center'})
    format_qty = book1.add_format({'num_format':'#,##0', 'align':'center'})
                                   
    sheet1 = writer.sheets['Calculations']
    sheet1.set_column('A:A',10)
    sheet1.set_column('B:B',30)
    sheet1.set_column('C:C',12)
    sheet1.set_column('D:D',45)
    sheet1.set_column('E:E',10,format_cost_decimals)
    sheet1.set_column('F:K',12,format_qty)
    sheet1.set_column('L:L',12,format_cost_decimals)
    sheet1.set_column('M:N',18,format_cost_decimals)
   
    writer.save()
    
    return df2


# Function to group results of mixed forecasts by category and opco

def calc_grouped():
    
    df3 = pd.read_excel(userfiles_dir + 'Cost Calculations.xlsx', index_col=[0,1,2,3], dtype={'part_number':str})
    df3_category = df3.groupby(by='part_category')['Total Cost','Rounded Total Cost'].sum()
    df3_opco = df3.groupby(by='opco')['Total Cost','Rounded Total Cost'].sum()
    return df3_category, df3_opco



# Function to generate quantity comparison sheet

def generate_qty_compare():
    
    df1 = df.groupby(['opco','part_category','part_number','description'])['unit_cost'].mean()
    df1 = pd.DataFrame(df1)
    
    df6 = df.loc[df['usage_weeks']==6]
    df6 = df6.groupby(['opco','part_category','part_number','description'])['expected_quantity'].sum()
    
    df13 = df.loc[df['usage_weeks']==13]
    df13 = df13.groupby(['opco','part_category','part_number','description'])['expected_quantity'].sum()
    
    df26 = df.loc[df['usage_weeks']==26]
    df26 = df26.groupby(['opco','part_category','part_number','description'])['expected_quantity'].sum()
    
    df52 = df.loc[df['usage_weeks']==52]
    df52 = df52.groupby(['opco','part_category','part_number','description'])['expected_quantity'].sum()
    
    df1['6-week Qty'] = df6
    df1['13-week Qty'] = df13
    df1['26-week Qty'] = df26
    df1['52-week Qty'] = df52
    df1['Chosen Qty'] = pd.Series()
    
    # Exporting comparative dataframe to excel
    
    writer = ExcelWriter(userfiles_dir + 'Quantity Comparison Sheet' +'.xlsx', engine='xlsxwriter')
    df1.to_excel(writer,'Comparison')
    
    book1 = writer.book
    format_qty = book1.add_format({'num_format':'#,##0', 'align':'center'})
    format_cost_decimals = book1.add_format({'num_format':'$ #,##0.00', 'align':'right'})
    format_text_to_number = book1.add_format({'num_format':'0'})
    
    sheet1 = writer.sheets['Comparison']
    sheet1.set_column('A:A',7.5)
    sheet1.set_column('B:B',30)
    sheet1.set_column('C:C',12,format_text_to_number)
    sheet1.set_column('D:D',50)
    sheet1.set_column('E:E',12,format_cost_decimals)
    sheet1.set_column('F:J',15,format_qty)
    
    writer.save()
    
    print('\nThe quantity comparison sheet will now open. Please review the numbers and choose appropriate quantities.\n')
    time.sleep(4)
    webbrowser.open(userfiles_dir+'Quantity Comparison Sheet.xlsx')


# Function to generate finance file
    
def generate_finance_file():
    
    finance_sheet1 = calc_mixed()
    del finance_sheet1['6-week Qty']
    del finance_sheet1['13-week Qty']
    del finance_sheet1['26-week Qty']
    del finance_sheet1['52-week Qty']
    del finance_sheet1['Rounded Unit Cost']
    finance_sheet1 = finance_sheet1.sort_index(ascending=False)
    finance_sheet2 = calc_grouped()[0]
    finance_sheet3 = calc_grouped()[1]
    
    writer = ExcelWriter(userfiles_dir + 'Finance File.xlsx', engine='xlsxwriter')
    finance_sheet1.to_excel(writer,'Projected Spend - By Part')
    finance_sheet2.to_excel(writer,'Projected Spend - By Category')
    finance_sheet3.to_excel(writer,'Projected Spend - By OpCo')
    
    book1 = writer.book
    format_cost_decimals = book1.add_format({'num_format':'$ #,##0.00', 'align':'right'})
    format_qty = book1.add_format({'num_format':'#,##0', 'align':'center'})
    format_text_to_number = book1.add_format({'num_format':'0'})
    
    sheet1 = writer.sheets['Projected Spend - By Part']
    sheet1.set_column('A:A',7.5)
    sheet1.set_column('B:B',30)
    sheet1.set_column('C:C',12)
    sheet1.set_column('D:D',50)
    sheet1.set_column('E:E',12,format_cost_decimals)
    sheet1.set_column('F:G',15,format_qty)
    sheet1.set_column('H:I',20,format_cost_decimals)
    
    sheet2 = writer.sheets['Projected Spend - By Category']
    sheet2.set_column('A:A',50)
    sheet2.set_column('B:C',20,format_cost_decimals)
    
    sheet3 = writer.sheets['Projected Spend - By OpCo']
    sheet3.set_column('A:A',20)
    sheet3.set_column('B:C',20,format_cost_decimals)
    
    label_format = book1.add_format()
    label_format.set_bold()
    label_format.set_align('center')
    
    total_format = book1.add_format()
    total_format.set_bold()
    total_format.set_align('right')
    total_format.set_num_format('$ #,##0.00')
        
    format_headers = book1.add_format()
    format_headers.set_bold()
    format_headers.set_align('center')
    
    sheet1.write('A1','Op Co.',format_headers)
    sheet1.write('B1','Part Category',format_headers)
    sheet1.write('C1','Part Number',format_headers)
    sheet1.write('D1','Description',format_headers)
    sheet1.write('E1','Unit Cost',format_headers)

    sheet2.write('A1','Part Category',format_headers)
                               
    sheet3.write('A1','Op Co.',format_headers)    
                            
    sheet3.write('A8','OVERALL',label_format)
    sheet3.write('B8','=SUM(B2:B7)',total_format)
    sheet3.write('C8','=SUM(C2:C7)',total_format)
 
    writer.save()
    
    print('\nThe finance file has been generated and will now open!\n')
    time.sleep(4)
    webbrowser.open(userfiles_dir+'Finance File.xlsx')


# Function for functionality selection
    
def choices():
    
    print('----------------------------------------------------------------------------------------------------------')
    
    while True:
        
        try:
            option = int(input('\nChoose function: \n1. Generate quantity comparison sheet \n2. Generate finance file \n3. Exit \n \n'))
    
            if option == 1:
                
                generate_qty_compare()
                do_more()
                break
                
            elif option == 2:
                
                generate_finance_file()
                do_more()
                break
                
            elif option == 3:
                
                print('\nThank you for using the app!\n')
                exit
                break
            
            else:
                print('\nInvalid, please try again.\n')
                continue
                                           
        except:
            
            print('\nInvalid, please try again.\n')
            continue
    
                
# Function to continue using the app
         
def do_more():
    
    while True:
        
        try:
            go_on = int(input('\nWould you like to continue using the app?\n 1. Yes\n 2. No\n'))
    
            if go_on == 1:
                
                choices()
                break
                
            elif go_on == 2:
                
                print('\nThank you for using the app!\n')
                exit
                break
                
            else:
                
                print('\nInvalid choice. Please try again.\n')
                continue
            
        except:
            print('\nInvalid choice. Please try again.\n')
            continue
  

#%% PROCESSING

   
# Get year for which to project

while True:
    
    try:
        year = int(input('\nPlease enter the desired year for the analysis:\n'))
    except:
        print('\nInvalid, please try again.\n')
        continue
    else:
        break
 
year_before = year - 1
year_after = year + 1

year_before = str(year_before)
year_after = str(year_after)

year_start = year_before + '-12-31'
year_end = year_after + '-01-01'

df = df.loc[df['expiration_date']<year_end]
df = df.loc[df['expiration_date']>year_start]

df = df.reset_index(drop=True)

choices()

