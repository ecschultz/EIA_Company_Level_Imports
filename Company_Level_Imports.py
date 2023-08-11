### Company Level Exports Refinery Info

# *** For Questions contact:  Ethan Schultz (ethan.schultz@conocophillips.com) ***
# EIA webpage: https://www.eia.gov/petroleum/imports/companylevel/archive/

#### Import Packages
#!pip install schedule
from bs4 import BeautifulSoup as bs
import requests
import pandas as pd
import win32com.client as win32
from datetime import datetime
import schedule
import time

DOMAIN = 'https://www.eia.gov'
URL = 'https://www.eia.gov/petroleum/imports/companylevel/archive/'
FILETYPE = '.xls'

def get_soup(url):
    return bs(requests.get(url).text, 'html.parser')

def get_links(page):
    result = page.find_all(class_="ico_xls")
    return result

page = get_soup(URL)
folder_links = get_links(page)
# print (folder_links)       
#######################################################################################

final_list = []

for file in folder_links:   #file_links
        filepath = file['href']
        if FILETYPE in filepath: 
            # print(DOMAIN + filepath) ###########
            final_list.append(DOMAIN + filepath)

#final_list

df = pd.DataFrame(final_list[:30], columns=['Links'])
##df = df.style.set_properties(subset=['Links'], **{'width-min': '300px'})
## apply style on the columns
##df.style.apply(lambda x: ["text-align:right"]*len(x))
#######################################################################################
df_list = []

for url in final_list:
    #data = requests.get(url)
    ##data_json = data.json()

    df_list.append(pd.read_excel(url))

    #df = pd.read_excel(url) #, sheet_name='IMPORTS'
    #df_list[-17:].append(df)

#quote_df = pd.concat(df_list[-17:])
final_df = pd.concat(df_list, ignore_index=True)
final_df['RPT_PERIOD'] =pd.to_datetime(final_df['RPT_PERIOD'], errors='coerce')

# Filter data between two dates
filtered_df = final_df.loc[(final_df['RPT_PERIOD'] >= '2017-01-01')]
filtered_df.sort_values(by='RPT_PERIOD',ascending=False)

# Drop any duplicates
filtered_df = filtered_df.drop_duplicates(subset=['RPT_PERIOD','R_S_NAME','LINE_NUM'],keep= 'last')

### Write the master file to the Publics folder ###
filtered_df.to_excel('S:\Commercial Assets\Comm Development (ADM095)\Market Analysis (ADM095)\PowerBI\EIA_Company_level_Imports_Info\company_level_imports.xlsx')

#######################################################################################
#######################################################################################

### Confirm email of script run

# datetime object containing current date and time
now = datetime.now()

# Change Date format dd/mm/YY H:M:S
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")


olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

# construct email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = "Company Level Imports (Refiners) Excel File has been updated"
mailItem.BodyFormat = 1
mailItem.Body = "Company Level Imports (Refiners) Excel File has been updated. \
                <br> \
                <br> \
                Refer to: S:\Commercial Assets\Comm Development (ADM095)\Market Analysis (ADM095)\Public Reports\RefineryInfo\Company_level_Imports_Info \
                <br> \
                <br> \
                To view updated PowerBI view: https://app.powerbi.com/groups/4d98a6f5-8af3-49e8-a85c-5e269e367a09/reports/59078b67-ba17-4664-9698-7919716f0a29"

mailItem.To = "ethan.schultz@conocophillips.com;Ishk.N.Varghese@conocophillips.com; Kaleb.P.Carr@conocophillips.com; federico.e.zamar@conocophillips.com"
mailItem.Sensitivity  = 2
# optional (account you want to use to send the email)
# mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('<email@gmail.com')))
mailItem.Display()
mailItem.Save()
mailItem.Send()
