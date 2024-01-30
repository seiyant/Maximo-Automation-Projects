# plannedHoursCorrection
# 
# This script intends to check work order planned hours and compare them to actual hours
#
# Author: Seiya Nozawa-Temchenko
##########################################################################################

# Load all relevant packages that were downloaded using pip
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from docx import Document
import re
import time
import datetime
import xlwings as xw

# Main function 
def main():
    start_date, end_date = date_ranger()

    browser = webdriver.Edge()

    print('Web browser initiated...\n')

    index, total = maximo_navigation(browser, start_date, end_date)

    extraction_excel(browser, index, total)

    print('We out')
    
    browser.quit()

def date_ranger():
    # Today's date
    todate = datetime.datetime.now() 
    print(todate)
    
    start_date = input('Enter start date as MM/DD/YYYY: ')
    
    while len(start_date) != 10:
        print('Enter the start date as MM/DD/YYYY and ensure it is a past date')
        start_date = input('Enter start date MM/DD/YYYY: ')

    date_format = '%m/%d/%Y'
    sdate = datetime.datetime.strptime(start_date, date_format)
    skey = str(sdate.month) + '/' + str(sdate.day) + '/' + str(sdate.year) + '/ 12:00 AM'

    today_innit =  input('Press y if end date is today: ')
    
    if today_innit == 'y':
        fdate = todate
        fkey = str(todate.month) + '/' + str(todate.day) + '/' + str(todate.year) + '/ 12:00 AM'
        
    else:
        print('End date is not today')
        end_date = input('Enter end date MM/DD/YYYY: ')
        while len(end_date) != 10:
            print('Enter the end date as MM/DD/YYYY and ensure it is a past date')
            start_date = input('Enter end date MM/DD/YYYY: ')
        fdate = datetime.datetime.strptime(end_date, date_format)
        fkey = str(fdate.month) + '/' + str(fdate.day) + '/' + str(fdate.year) + '/ 12:00 AM'

    dtime = fdate.date() - sdate.date()
    print(f"Your given time frame is {dtime.days} days")

    return skey, fkey

def maximo_navigation(browser, start_date, end_date):
    wait = WebDriverWait(browser, 20)
    browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
    browser.maximize_window()

    # Extract login information from text file
    credentials = {}
    with open('config.txt', 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            credentials[key] = value
            
    # Enter login information
    UserElem = wait.until(EC.element_to_be_clickable((By.ID, 'username')))
    UserElem.send_keys(credentials['username'])

    Cont1Elem = browser.find_element(By.XPATH, '/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button')
    Cont1Elem.click()

    passElem = wait.until(EC.element_to_be_clickable((By.ID, 'password')))
    passElem.send_keys(credentials['password'])

    Cont2Elem = browser.find_element(By.XPATH, '/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button')
    Cont2Elem.click()

    # Navigate to Work Order Tracking
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[7]/iframe')))
    browser.switch_to.frame(iframe)
    wotrackElem = wait.until(EC.element_to_be_clickable((By.ID, 'FavoriteApp_WOTRACK')))
    wotrackElem.click()

    # Navigate to More Search Fields
    triple_dot = wait.until(EC.element_to_be_clickable((By.ID, 'quicksearchQSMenuImage')))
    triple_dot.click()
    more_search = wait.until(EC.element_to_be_clickable((By.ID, 'menu0_SEARCHMORE_OPTION_a')))
    more_search.click()

    # Set Status to CLOSE, FINISHED, or WAITCLOSE
    status = wait.until(EC.element_to_be_clickable((By.ID, 'm449c436f-tb')))
    status.click()
    status.send_keys('=CLOSE,=FINISHED,=WAITCLOSE')

    # Garbage value to ensure things load
    garbage_value = wait.until(EC.element_to_be_clickable((By.ID, 'm8db33e5c-tb')))
    garbage_value.click()

    # Set History? to Y
    history = wait.until(EC.element_to_be_clickable((By.ID, 'mdd9512d5-tb')))
    history.click()
    history.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    history.send_keys('Y')
    garbage_value.click()
    
    # Set sched Start
    startget = wait.until(EC.element_to_be_clickable((By.ID, 'mafd0ceda-tb')))
    startget.click()
    startget.send_keys(start_date)    
    garbage_value.click()

    # Set Target Finish
    endget = wait.until(EC.element_to_be_clickable((By.ID, 'mdfba3a55-tb')))
    endget.click()
    endget.send_keys(end_date)
    findbutton = wait.until(EC.element_to_be_clickable((By.ID, 'm4fd840b0-pb')))
    findbutton.click() 
    
    # Find number of Work Orders
    time.sleep(2) # Magic number
    wostring = wait.until(EC.element_to_be_clickable((By.ID, 'm6a7dfd2f-lb3')))
    match = re.search(r'\((\d+) - (\d+) of (\d+)\)', wostring.text)

    if match:
        i = int(match.group(1))
        max = int(match.group(3))
    else:
        i = 0
        max = 0
        print("No results in this time frame. If issue persists change the sleep length")

    print(f'Total Work Orders: {max}')
    return i, max

def extraction_excel(browser, index, total):
    wait = WebDriverWait(browser, 20)

    # Cycle through each line and page 
    while i <= max:
        # line  1: m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]
        # line  2: m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:1]
        # line 20: m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:19]

        # line  1: m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:20]
        # line 20: m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:39]
    category, craft, process cond, desc, job plan #, planned hrs, new plan?, actusal hours, laborers
    

main()
#Go to WO Tracking
#Go to More Search Fields
#Set History to 'Y'
#Set Target Start to First of a month
#Set Target Finish to Last day of a month
# Target dates can be manually entered
#Normal distribution of difference in work hours
#Job plan copy and search again
#Plans contain hours
#Actual contains real work hours done

#CONSIDERATIONS
#Manage repeats, if information is identical
