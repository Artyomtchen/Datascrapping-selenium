# -*- coding: utf-8 -*-
"""
Created on Thu May  7 16:40:32 2020

Web crawler importing monthly Passenger and light truck vehicle sales in Norway by model (big data) using selenium and nested loops 

@author: artyom
"""
# =============================================================================
# Modules
# =============================================================================
import traceback
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os
import glob
import re

# =============================================================================
# Webdriver and loging in
# =============================================================================

#Set downloadpath to a desired workwolder

chrome_options = webdriver.ChromeOptions()
downloadpath=r'C:\Users\artyom\Desktop\python\OVF data collection\Collected data'
prefs = {'download.default_directory' : downloadpath}
chrome_options.add_experimental_option('prefs', prefs)
browser = webdriver.Chrome(chrome_options=chrome_options)


#start the webdriver, navigate to the page and log in to page
# browser = webdriver.Chrome(chrome_options=options)

url='https://statistikk.ofv.no/'
browser.get(url) 
username = browser.find_element_by_name('dirty_username') #username form field
password = browser.find_element_by_name("dirty_password") #password form field
username.send_keys("here username used on the webpage")
password.send_keys("password used on the webpage")
submitButton = browser.find_element_by_name("submit")
submitButton.click()

#Selectors of vehicle type and fuel type (similar to dropdown menues found on webpages)

cartype = ['Personbiler', r'Vare / Kombi < 3.5t'] 
fueltype =['Bensin', 'Diesel', 'Gass', 'Elektrisitet', 'Hydrogen', 'Hybrid', 'Plugin Hybrid']

#Telling the webdriver that data will be collected by vehicle model (another option is to collect data by brand)
browser.get('https://statistikk.ofv.no/secure/lists/week/normal/menu.asp') #navigate to the correct site
model_select = Select(browser.find_element_by_id('type'))
model_select.select_by_visible_text('Modellfordelt')

#Get a list of the names of the period we want to look at. This code downloads all periods starting from Jan 2005
#If only latest data should be collected, the code should include the line periodtype=periotype[0] 

listofelements=browser.find_elements(By.XPATH,'//*[@name="periode"]/option')
periodtype=[item.text for item in listofelements[:len(listofelements)-60]]#only include time 2005-2019 because of error

#check what value corresponds to listofelements-59 - this is the last period we collect from the website
#check what value corresponds to the latest period we collect the data for 
listofelements[len(listofelements)-59].text
listofelements[0].text

# set download folder location
# set list_all_files that will be later used for checking of whether new data has appread on the website and should be collected
os.chdir(downloadpath)
list_all_files = glob.glob('*.xls')

# =============================================================================
# Nested loop 
# =============================================================================

# Loop through all options from dropdown menu and download corresponding files into the download folder on the server
# If a loop inspection mode is desired (just to see how loop is running) we need to a) comment out lines that click through the data
#b) and set sleeptime to 0 c) uncomment ifs and breaks d) and run the nested loop
#try except operators make nested loop print and error if exception occurs, otherwise continue looping


try:
    for car in cartype:
        car_select = Select(browser.find_element_by_id('cartype'))
        car_select.select_by_visible_text(car)
        for fuel in fueltype:
            fuel_select = Select(browser.find_element_by_id('sel_fuel'))
            fuel_select.select_by_visible_text(fuel)
            for index, period in enumerate(periodtype):
                filename = period+'-'+fuel+'-'+car+'.xls'
                filename = re.sub(r'[\\/*?:"<>|]',"_",filename)
                if filename in list_all_files :
                    print( '%s Already collected' % (filename) )
                    break
                select = Select(browser.find_element_by_id('periode'))
                select.select_by_visible_text(period)
                elem=browser.find_element_by_xpath('//*[@id="menufrm"]/table[1]/tbody/tr[2]/td[3]/a')
                elem.click()
                time.sleep(2)
                print('%s %s (%i / %i), %s' % ( car, fuel, index+1, len(listofelements)-60, period))
                latest_file  = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)[-1]
                print(filename)            
                os.rename(latest_file, filename)
                print(latest_file )
                # if period == 'Mars - 2020':
                #     break
except:
    print('Error with %s %s (%i / %i), %s' % ( car, fuel, index+1, len(listofelements)-60, period))
    
#If the "done" message gets printed at the end, the download was succsessful            
print('done')

# =============================================================================
# Collect trucks and buses
# =============================================================================

#Selectors of vehicle type and fuel type (similar to dropdown menues found on webpages)
cartype = ['Lastebiler', 'Busser'] # a list that does not include combination

#Selectors of vehicle type and fuel type (similar to dropdown menues found on webpages)
browser.get('https://statistikk.ofv.no/secure/lists/week/normal/menu.asp') #navigate to the correct site

#Get a list of the names of the period we want to look at. This code downloads all periods starting from Jan 2005
#If only latest data should be collected, the code should include the line periodtype=periotype[0]

listofelements=browser.find_elements(By.XPATH,'//*[@name="periode"]/option')
periodtype=[item.text for item in listofelements[:len(listofelements)-60]]#only include time 2005-2019 because of error

os.chdir(downloadpath)
list_all_files = glob.glob('*.xls')

# =============================================================================
# Nested loop 
# =============================================================================

# Loop through all options from dropdown menu and download corresponding files into the download folder on the server
# If a loop inspection mode is desired (just to see how loop is running) we need to a) comment out lines that click through the data
#b) and set sleeptime to 0 c) uncomment ifs and breaks d) and run the nested loop
#try except operators make nested loop print and error if exception occurs, otherwise continue looping
        
try:
    for car in cartype:
        car_select = Select(browser.find_element_by_id('cartype'))
        car_select.select_by_visible_text(car)
        for index, period in enumerate(periodtype):
            filename = period+'-'+car+'.xls'
            filename = re.sub(r'[\\/*?:"<>|]',"_",filename)
            if filename in list_all_files :
                    print( '%s Already collected' % (filename) )
                    break
            select = Select(browser.find_element_by_id('periode'))
            select.select_by_visible_text(period)
            elem=browser.find_element_by_xpath('//*[@id="menufrm"]/table[1]/tbody/tr[2]/td[3]/a')
            elem.click()
            time.sleep(2)
            print('%s, %i , %s' % ( car, index+1, period))
            latest_file  = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)[-1]
            print(filename)            
            os.rename(latest_file, filename)
            print(latest_file )
            # if period == 'Januar - 2020':
            #     break

        # if car == 'Lastebiler':
            # break
except:
    print('Error with %s, %i, %s' % ( car, index+1, period))
    
    
print('done')

# =============================================================================
#pandas dataframe that gethers cleans and writes data to Excel
# =============================================================================

#Define correct directory where all uploaded metafiles from OVF are located . By default it is Download folder
# #The directory should be empty before downloading any files.
# This is the path to the  directory folder. 
downloadpath=r'C:\Users\artyom\Desktop\python\OVF data collection\Collected data'
os.chdir(downloadpath)
#Collect all the names of the files in the directory and put it in a list called "directory"
directory = os.listdir(downloadpath)

#Make a dataframe where all data will be collected
df_output=pd.DataFrame()
liste=[]
n=0
#loop through the list of file names in the download folder
for file in directory:
    if n % 500==1:
        print(n)
    #read the files in the directory one by one
    df = pd.read_html(file)

    #Check format the file has so that the formating will be correct
    if len(df)>1 and len(df[0])==10: #check that the file is correct
        # allocate correct data to variables
        date=df[0][0][0]
        cartype=df[0][0][1]
        fueltype=df[0][0][3]
        month, year = date.split('-', 1 )
        month=month.replace('="','').strip()
        year=year.replace('"','').strip()

        # insert variables into respective columns in the df
        df=df[1]
        df['Car']=cartype
        df['Fuel']=fueltype
        df['Month']=month
        df['Year'] =year
        
        df_output = df_output.append(df)     
        n+1

    elif len(df)>1 and len(df[0])==8:
        # do the same as above for trucks and buses
        date=df[0][0][0]
        cartype=df[0][0][1]
        fueltype='Not specified'
        month, year = date.split('-', 1 )
        month=month.replace('="','').strip()
        year=year.replace('"','').strip()

        df=df[1]
        df['Car']=cartype
        df['Fuel']=fueltype
        df['Month']=month
        df['Year'] =year

        df_output = df_output.append(df)
        time.sleep(1)
        n+1



    elif len(df)==1:
        liste.append(file)
        n+1

    else:   
        print(file)
        n+1
        continue
       
df_final=df_output

#mapping Norwegian month names ('Mars') to correct month values (3):
    
months2numbers = [['Januar',1],
                  ['Februar',2],
                  ['Mars',3],
                  ['April',4],
                  ['Mai',5],
                  ['Juni',6],
                  ['Juli',7],
                  ['August',8],
                  ['September',9],
                  ['Oktober',10],
                  ['November',11],
                  ['Desember',12]]
                  
df_month_mapping = pd.DataFrame(months2numbers, columns=['Month','MonthNum'])
df_final = pd.merge(df_final, df_month_mapping, on=['Month'], how='left')

df_final['Year']=df_final['Year'].astype(int)



#creating a matching spreadsheet for clean fuel names and vehicle categories
df_check=df_final['Fuel'].drop_duplicates().reset_index()
df_check['Fuelclean']=df_check['Fuel']
#cleaning fuels of spaces in front and after names as well as = and ' signs
df_check['Fuelclean']=df_check['Fuelclean'].str.lstrip('="')
df_check['Fuelclean']=df_check['Fuelclean'].str.rstrip('filter"')
df_check['Fuelclean']=df_check['Fuelclean'].str.lstrip()
df_check['Fuelclean']=df_check['Fuelclean'].str.rstrip()

#cleaning car types of spaces in front and after names as well as = and ' signs
df_check2=df_final['Car'].drop_duplicates().reset_index()
df_check2['Carclean']=df_check2['Car']
df_check2['Carclean']=df_check2['Carclean'].str.lstrip('="')
df_check2['Carclean']=df_check2['Carclean'].str.rstrip('"')
df_check2['Carclean']=df_check2['Carclean'].str.lstrip()
df_check2['Carclean']=df_check2['Carclean'].str.rstrip()

#inserting back into final dataframe cleaned fuel name and vehicle category data
df_final = pd.merge(df_final, df_check, on=['Fuel'], how='left')
df_final = pd.merge(df_final, df_check2, on=['Car'], how='left')

#renaming first column from numerical value to "Modelnamedirty"
df_final.rename(columns={df_final.columns[0]:'Modelnamedirty'}, inplace=True)


#cleaning model names: removing 1. at the begining and stripping of any spaces in front and behind the data
df_cleanmodels=df_final.filter(['Modelnamedirty'], axis=1).drop_duplicates()
df_cleanmodels['Modelnameclean']=df_cleanmodels['Modelnamedirty'].str.lstrip("0123456789.")
df_cleanmodels['Modelnameclean']=df_cleanmodels['Modelnameclean'].str.lstrip()
df_cleanmodels['Modelnameclean']=df_cleanmodels['Modelnameclean'].str.rstrip()
df_cleanmodels.reset_index()


#Merging cleaned model names back into the initial dataframe
df_final = pd.merge(df_final, df_cleanmodels, on=['Modelnamedirty'], how='left')

df_final.drop(df_final.columns[[0,16,18]], axis=1, inplace=True)
df_final.drop(df_final.iloc[:,1:10], axis=1, inplace=True)

#renaming first column from numerical value to "Modelnamedirty" once again
df_final.rename(columns={df_final.columns[0]:'Quantity'}, inplace=True)


#When all files are looped thwoug the file merged.xlsx is created in the download folder and the
#merge is completed. The file will not be created if an error occures.
os.chdir(r'C:\Users\artyom\Desktop\python\OVF data collection')
df_final.to_excel('merged.xlsx')
