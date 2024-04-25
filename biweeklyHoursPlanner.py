# biweeklyHoursPlanner.py
# 
# This script intends to gather work orders for the next 2 weeks to find the total work work hours
#
# Author: Ishmam Raza Dewan, Seiya Nozawa-Temchenko
##########################################################################################

#Load all relevant packages that were downloaded using pip
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import datetime, time
import xlwings as xw

browser = webdriver.Edge() 

wb = xw.Book(r'P:\All\SERVER REPORTS\2 Week Plan.xlsx') #excel workbook to be used
sheet = xw.sheets['April 15'] #increase sheet number biweekly, or hardcode name

#Changing what this script clicks on requires your browser dev tools
#Each object has an ID, inspect element to hover over object, click to find ID

actions = ActionChains(browser) 
wait = WebDriverWait(browser, 20)
browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
browser.maximize_window()

# Extract login information
credentials = {}
with open('P:\All\Engineering\Projects\Python Scripts\Seiya SEP 2023-APR 2024\Maximo-Automation-Projects\config.txt', 'r') as file:
    for line in file:
        key, value = line.strip().split('=')
        credentials[key] = value

# Enter login information
UserElem = wait.until(EC.element_to_be_clickable((By.ID, "username")))
UserElem.send_keys(credentials['username'])

Cont1Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
Cont1Elem.click()

passElem = wait.until(EC.element_to_be_clickable((By.ID, "password")))
passElem.send_keys(credentials['password'])

Cont2Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
Cont2Elem.click()

# Navigate to Work Order Tracking
iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[7]/iframe")))
browser.switch_to.frame(iframe)
wotrackElem = wait.until(EC.element_to_be_clickable((By.ID, "FavoriteApp_WOTRACK")))
actions.move_to_element_with_offset(wotrackElem, 5, 5).click().perform()

time.sleep(3)
###############################################################################
dt = datetime.datetime.now() #current time
delta = datetime.timedelta(days = 14) #time difference
ft = dt + delta #14 days fromt oday

smonth = str(dt.month) #month
sday = str(dt.day) #day
syear = str(dt.year) #year

fmonth = str(ft.month) #month 14 days from today
fday = str(ft.day) #day 14 days from today
fyear = str(ft.year) #year 14 days from today

# Navigate to More Search Fields
triple_dot = wait.until(EC.element_to_be_clickable((By.ID, 'quicksearchQSMenuImage')))
triple_dot.click()
more_search = wait.until(EC.element_to_be_clickable((By.ID, 'menu0_SEARCHMORE_OPTION_a')))
more_search.click()

# Set Types to HKG, INR, or PPM
types = wait.until(EC.element_to_be_clickable((By.ID, 'med325893-tb')))
types.click()
types.send_keys('=HKG,=INR,=PPM') #ishmam had this idk why its set to these and not all

# Garbage value to ensure things load
garbage_value = wait.until(EC.element_to_be_clickable((By.ID, 'm8db33e5c-tb')))
garbage_value.click()

# Set Status to CLOSE, FINISHED, or WAITCLOSE
status = wait.until(EC.element_to_be_clickable((By.ID, 'm449c436f-tb')))
status.click()
status.send_keys('=RELEASED,=WPLAN,=WSCHED,=WKIT')
garbage_value.click()

# Set Target Start
startget = wait.until(EC.element_to_be_clickable((By.ID, 'm3cdc438b-tb')))
startget.click()
startget.send_keys(smonth + '/' + sday + '/' + syear + ' 12:00 AM')    
garbage_value.click()

# Set Final Target Start
finalget = wait.until(EC.element_to_be_clickable((By.ID, 'mac635e1a-tb')))
finalget.click()
finalget.send_keys(fmonth +'/' + fday + '/' +fyear + ' 12:00 AM')
findbutton = wait.until(EC.element_to_be_clickable((By.ID, 'm4fd840b0-pb')))
findbutton.click()

time.sleep(5)
wostring = wait.until(EC.element_to_be_clickable((By.ID, 'm6a7dfd2f-lb3')))
match = re.search(r'\((\d+) - (\d+) of (\d+)\)', wostring.text)

if match:
    numberofWOs = int(match.group(3))
else:
    numberofWOs = 0
    print("No results found in this time frame")

print(f'Total Work Orders: {numberofWOs}')

i = 0
while i < numberofWOs:
    # Navigate to Work Order
    if i == 0:
        wo_html = f'm6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:{i}]'
        findwo = wait.until(EC.element_to_be_clickable((By.ID, wo_html)))
        findwo.click()
        time.sleep(2)
    else:
        try:
            findwo = wait.until(EC.element_to_be_clickable((By.ID, 'toolactions_NEXT-tbb_anchor')))
            findwo.click()
            time.sleep(2)
            nav2wo = wait.until(EC.element_to_be_clickable((By.ID,'mbf28cd64-tab_anchor')))
            nav2wo.click()

        except: #system message
            syserr = wait.until(EC.element_to_be_clickable((By.ID, 'm96ad0396-pb')))
            syserr.click()
            time.sleep(2)
            findwo = wait.until(EC.element_to_be_clickable((By.ID, 'toolactions_NEXT-tbb_anchor')))
            findwo.click()
            time.sleep(2)
            nav2wo = wait.until(EC.element_to_be_clickable((By.ID,'mbf28cd64-tab_anchor')))
            nav2wo.click()
            time.sleep(2)

    # Navigate in Work Order
    WO = wait.until(EC.element_to_be_clickable((By.ID, 'm52945e17-tb'))).get_attribute('value')
    print(f'Work Order: {WO}')
    description = wait.until(EC.element_to_be_clickable((By.ID, 'md42b94ac-tb'))).get_attribute('value')
    print(f'Description: {description}')
    jobplan = wait.until(EC.element_to_be_clickable((By.ID, 'mfe7bb84-tb'))).get_attribute('value')
    print(f'Job Plan: {jobplan}')
    tstart = wait.until(EC.element_to_be_clickable((By.ID, 'm651c06b0-tb'))).get_attribute('value')
    print(f'Target Start: {tstart}')
    asset = wait.until(EC.element_to_be_clickable((By.ID, 'm3b6a207f-tb'))).get_attribute('value') 
    print(f'Asset: {asset}')

    # Navigate to Plans
    plans = wait.until(EC.element_to_be_clickable((By.ID, 'm356798d1-tab_anchor')))
    plans.click()
    time.sleep(5)

    # There may be multiple laboreres
    labstring = wait.until(EC.element_to_be_clickable((By.ID, 'm5e4b62f0-lb3')))
    match = re.search(r'\((\d+) - (\d+) of (\d+)\)', labstring.text)
    labor_max = int(match.group(3))
    j = 0; plannedhrs = 0

    while j < labor_max:
        time.sleep(2)
        print(f'Index {j + 1} of {labor_max}')

        hours_html = f'm5e4b62f0_tdrow_[C:9]_txt-tb[R:{j}]'
        hrs, mins = map(int, wait.until(EC.element_to_be_clickable((By.ID, hours_html))).get_attribute("value").split(':'))

        plannedhrs = hrs + mins / 60 + plannedhrs
        print(f'Planned hours: {plannedhrs}')

        if j % 6 == 5:
            pej_flipa = wait.until(EC.element_to_be_clickable((By.ID, 'm5e4b62f0-ti7')))
            pej_flipa.click()
            time.sleep(3)
        
        j += 1
        
    # Excel writing
    sheet['A' + str(5+i)].value = WO 
    sheet['B' + str(5+i)].value = jobplan
    sheet['C' + str(5+i)].value = description
    sheet['E' + str(5+i)].value = plannedhrs
    sheet['D' + str(5+i)].value = tstart
    sheet['F' + str(5+i)].value = asset

    if sheet['E' + str(5+i)].value == 0:
        sheet['E' + str(5+i)].value = 0.5 

    print(f'Work Order {WO} completed...')
    i += 1

wb.save() 
browser.quit()
