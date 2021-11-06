"""
Created on Fri Jun  4 11:45:26 2021
@author: Kenneth Liang

This script pulls all the required CSV files from a specific website that need to be compiled together

Packages that need to be installed
pip install requests
pip install beautifulsoup4
pip install lxml
pip install pandas
pip install io
pip install PySimpleGUI
pip install pyinstaller

virtualenv venv
venv\Scripts\activate

path\to\pyinstaller.exe --onefile --paths path\to\venv\Lib\site-packages file.py
pyinstaller --onefile -w 'Install_Base_Pull.py'

"""
# import libraries
import requests # for making website data requests
from requests.auth import HTTPBasicAuth

from bs4 import BeautifulSoup # for interpreting HTML data

import pandas as pd
import io

# For GUI
import PySimpleGUI as sg

layout = [
    [sg.Text('Please enter your Windows user name and password')],
    [sg.Text('User Name', size =(15, 1)), sg.InputText(password_char='*')],
    [sg.Text('Password', size =(15, 1)), sg.InputText(password_char='*')],
    [sg.Submit(), sg.Cancel()]
]
    
window = sg.Window('Credentials', layout)
event, values = window.read()
window.close()

# check if user closed or cancelled the password entry window
if event == sg.WIN_CLOSED or event == 'Cancel':
    sg.popup('Script stopped by user.')
    pass
else:
    
    user_name = values[0]
    password = values[1]
    
    try:
        # Make a get request and get the list of parts
        response = requests.get('ENTER WEBSITE LINK TO BE SCRAPED', auth = HTTPBasicAuth(user_name, password), verify=False)
    except:
        sg.popup('Connection error. Script stopped')
        pass
    else:
        # Check to see if password and username were authenticated correctly by the website
        if response.status_code == 401:
            sg.popup('Your user name and password are not correct. Script stopped')
            pass
        else:
            # Message to user that the script is running
            layout_msg = [[sg.Text('Script is running please wait')]]
            # Create the window
            window_msg = sg.Window('',layout_msg,keep_on_top=True)
            # Launch message window
            event, values = window_msg.read(timeout=1)
            
            # BeautifulSoup web content interpreter
            soup = BeautifulSoup(response.text,'lxml')
    
            parts = [] # initialize empty list
    
            # List of parts
            for table in soup.find_all('table'):
    
                for a in table.find_all('a'):
                    parts.append(a.text)
    
            # Intialize Pandas Dataframe
            df_csv = pd.DataFrame()
    
            # Make a get request and get the CSV file. Download into virtual memory instead of as a file
            for part in parts[1:]:
                
                CSV_URL = f'ENTER WEBSITE LINK TO BE WHERE CSV FILE CAN BE PULLED {part}'
    
                try:
                    with requests.Session() as s:
                        download = s.get(CSV_URL, auth = HTTPBasicAuth(user_name, password), verify=False)
                        
                        # utf-8 is the CSV format
                        decoded_content = download.content.decode('utf-8')
                        
                        # read into pandas dataframe
                        data = pd.read_csv(io.StringIO(decoded_content), delimiter=',',low_memory=False, dtype=str)
                        df_csv = df_csv.append(data)
                except:
                    pass
            
            # save dataframe as a CSV file
            df_csv.to_csv('network_hwdb.csv',encoding='utf-8',index=False)
            window_msg.close()