'''
CREATED BY: KENNETH LIANG
DATE: JULY, 2021

This script scrapes the Moxie website for instructors who have upcoming classes. It collects the instructor name, their Moxie page and Instagram handle
'''

# Chrome browser control
from selenium import webdriver
from selenium.webdriver import ChromeOptions

import os
import time

import re

import requests
from bs4 import BeautifulSoup
from xlsxwriter import Workbook

# remove duplicates from list
def remove_duplicates(seq):
    seen = set()
    seen_add = seen.add
    return [x for x in seq if not (x in seen or seen_add(x))]

# Open a chrome browser window, this needs to use the chrome driver exe
options = ChromeOptions() 
#options.headless = True
browser = webdriver.Chrome(executable_path=os.getcwd()+'\\chromedriver.exe',options=options)
browser.implicitly_wait(2)

# open chrome browswer and finite scroll to the bottom
browser.get('https://moxie.xyz/classes')
browser.maximize_window()

# Get scroll height after first time page load
last_height = browser.execute_script("return document.body.scrollHeight")

# Scroll while not at the bottom of the page. The Moxie website uses an infinity scroll page
while True:
    
    # Scroll down to bottom
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    
    # Wait to load page
    time.sleep(2)
    
    # Calculate new scroll height and compare with last scroll height
    new_height = browser.execute_script("return document.body.scrollHeight")
    
    if new_height == last_height:
        break
    
    last_height = new_height
    
# open text file of web page
textfile = open('...\\Moxie.txt', mode='r', encoding="utf8")
filetext = textfile.read()
textfile.close()

# find text with the following pattern which identifies the instructors. /_____& within the blank space is the instructors Moxie link
matches = re.findall("'/\w+'", filetext)

unique_matches = remove_duplicates(matches)

moxie_details = []

for instructor in map(lambda unique_matches:unique_matches.strip("'"),unique_matches):
    
    user_name, instagram = "",""
    
    response = requests.get(f'https://moxie.xyz{instructor}')
    soup = BeautifulSoup(response.text,'lxml')
    
    for span in soup.find_all('span', attrs={'class':'text-truncate profile-display-name'}):
        user_name = span.text
    
    for div in soup.find_all('div', attrs={'class':'text-truncate text-center text-md-left font-size-28 profile-display-name'}):
        user_name = div.text
    
    for div in soup.find_all('div', attrs={'class':'instructor-username d-flex align-items-center mr-3 mr-md-0 mb-md-2'}):
        instagram = div.find('a')['href']
    
    for div in soup.find_all('div', attrs={'class':'instructor-username d-flex align-items-center mb-md-2'}):
        instagram = div.find('a')['href']
        
    moxie_details.append([user_name, instructor, instagram])

# Output results of web scraping to a new excel file
with Workbook('Moxie_Details.xlsx') as newworkbook:
    worksheet = newworkbook.add_worksheet()

    # First row are the header names
    worksheet.write_row(0,0,['Name','Moxie Link','Instagram Link'])

    # Output the row data
    for row_num, data in enumerate(moxie_details):
        worksheet.write_row(row_num+1, 0, data, newworkbook.add_format({'text_wrap':True, 'valign':'vcenter'}))