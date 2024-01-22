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
import datetime, time
import xlwings as xw

browser = webdriver.Edge() 

wb = xw.Book(r'\\igashfs1\shared\All\SERVER REPORTS\2 Week Plan.xlsx') #excel workbook to be used
sheet = xw.sheets[0] #first sheet

#Changing what this script clicks on requires your browser dev tools
#Each object has an ID, inspect element to hover over object, click to find ID

actions = ActionChains(browser) 
wait = WebDriverWait(browser, 20)
browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
browser.maximize_window()

# Extract login information
credentials = {}
with open('P:\All\Engineering\Seiya Nozawa-Temchenko\Maximo Automation\Maximo-Automation-Projects\config.txt', 'r') as file:
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
time.sleep(1)

searchElem = browser.find_element(By.XPATH, "//*[@id='quicksearch_QSButtonDiv']/div[2]") #search button
actions.move_to_element(searchElem)
time.sleep(1)
actions.click()
actions.perform()
time.sleep(2)

search2Elem = browser.find_element(By.XPATH, "//*[@id='quicksearch_QSButtonDiv']/div[2]") #second search button
actions.move_to_element_with_offset(search2Elem,0,50)
time.sleep(1)
actions.click()
actions.perform()
time.sleep(5)

try:
    typeElem = browser.find_element(By.ID, "med325893-tb") #type text box. If find by ID fails it will find using full xpath
except:###NEW
    typeElem = browser.find_element(By.XPATH, "/html/body/form/div/table[3]/tbody/tr/td[3]/table/tbody/tr[2]/th/table/tbody/tr/td/table[2]/tbody/tr/td/div/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr/td[3]/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/input")
    
typeElem.send_keys('=HKG,=INR,=PPM') #types
time.sleep(3)

statusElem = browser.find_element(By.ID, "m449c436f-tb") #status textbox
statusElem.send_keys('=RELEASED,=WPLAN,=WSCHED') #status types
time.sleep(2)


startElem = browser.find_element(By.ID, "m3cdc438b-tb") #start time text box
startElem.send_keys(smonth + '/' + sday + '/' + syear + ' 12:00 AM') #send start time using the times gathered earlier
time.sleep(3)
finElem = browser.find_element(By.ID, "mac635e1a-tb") #final start time textbox
time.sleep(2)
finElem.click()
time.sleep(1)
finElem.send_keys(fmonth +'/' + fday + '/' +fyear + ' 12:00 AM') #send final start time using the times gathered earlier
time.sleep(2)


findElem = browser.find_element(By.ID, "m4fd840b0-pb") #find button
findElem.click()
time.sleep(10)
numberofWOs = browser.find_element(By.ID, "m6a7dfd2f-lb3") #text that says number of WOs
numberofWOs = numberofWOs.text #saves text in variable
print(numberofWOs)

if len(numberofWOs) == 14: #converts the whole string into just the number
    numberofWOs = numberofWOs[11] + numberofWOs[12] #if the string is 14 characters long, only take the 11th and 12th character
else:
    numberofWOs = numberofWOs[11] + numberofWOs[12] + numberofWOs[13] #otherwise use 11, 12, and 13 (3 figure number of work orders)

numberofWOs = int(numberofWOs) #convert to integer


time.sleep(2)
wo1Elem = browser.find_element(By.ID, "m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]") #change ID here if diff
type(wo1Elem) #wo1 = work order one
time.sleep(1)
wo1Elem.click() #click first work order
time.sleep(1)
i=0
while i<numberofWOs:
    time.sleep(4)
    wb.save() #save in excel
    time.sleep(2)
    
    WO = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/input").get_attribute("value");
    description = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr[3]/td/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/input").get_attribute("value");
    jobplan = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/input").get_attribute("value");
    tstart = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/input").get_attribute("value");
    asset = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/table/tbody/tr[3]/td/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td/div/table/tbody/tr/td/table/tbody/tr[10]/td/input[1]").get_attribute("value");
    PlansElem = browser.find_element(By.ID, "m356798d1-tab_anchor") #all of the above are text boxes, the .get_attribute() automatically takes the value in a single line. The left is a tab button
    time.sleep(3)
    type(PlansElem)
    PlansElem.click()#click plans tab
    time.sleep(10) 
    try:
        plannedhrs1 = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[11]/input").get_attribute("value");
    except:
        plannedhrs1 = 'no data' #try to find the first planned hours. If none, will say no data

    if plannedhrs1 == 'no data':
        time.sleep(2)
        try: #try again just to be sure
            plannedhrs1 = browser.find_element(By.XPATH, "/html/body/form/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[11]/input").get_attribute("value");
        except:
            print('even on a 2nd try plannedhrs1 has no data. Maybe there really is no data')

    try:
        plannedhrs = int(plannedhrs1[0])
    except:
        print('no data') #plannedhrs is the sum variable of plannedhrs1, plannedhrs2, and plannedhrs3. It changes depending on work order so I just trial and error it here. If the line above fails it just prints no data.

    try:
        plannedhrs = int(plannedhrs1[0] + plannedhrs1[1])
    except:
        print('no double digits or still no data') 
        
    if plannedhrs1 == 'no data':
        plannedhrs = 'no data'
        
#from the above, if no data in plannedhrs1, then plannedhrs is no data. otherwise, it will test with 1
#character, and then test again with 2.

#plannedhrs 2 section
        
    try:
        plannedhrstwo = browser.find_element(By.ID, "m5e4b62f0_tdrow_[C:9]_txt-tb[R:1]").get_attribute("value");
    except:
        plannedhrstwo = 'no data'
    else:
        plannedhrs2 = plannedhrstwo[0]

    try:
        plannedhrs2 = int(plannedhrs2)
    except:
        print('no data')

    try:
        plannedhrs2 = int(plannedhrstwo[0] + plannedhrstwo[1])
    except:
        print('less than 10 hours or still no data')

        
    try:
        plannedhrs = plannedhrs + plannedhrs2 #if plannedhrs2 doesnt exist, it prints no data for planned hrs 2
    except:
        print('no data for plannedhrs2')


#plannedhrs 3 section (identical to plannedhrs 2)

    try:
        plannedhrsthree = browser.find_element(By.ID, "m5e4b62f0_tdrow_[C:9]_txt-tb[R:2]").get_attribute("value");
    except:
        plannedhrsthree = 'no data'
    else:
        plannedhrs3 = plannedhrsthree[0]

    try:
        plannedhrs3 = int(plannedhrs3)
    except:
        print('no data')

    try:
        plannedhrs3 = int(plannedhrsthree[0] + plannedhrsthree[1])
    except:
        print('less than 10 hours or still no data')
        
    try:
        plannedhrs = plannedhrs + plannedhrs3
    except:
        print('no data for plannedhrs3')

#if all of these fail, then from the plannedhrs1 section, plannedhrs will equal 'no data'
        
#write to excel and move on to next wo
    sheet['A' + str(5+i)].value = WO #print WO value to cell A:5+i. i is the number of times the loop has run, which counts the work order number. the 5 added is because of the formatting of the excel sheet.
    sheet['B' + str(5+i)].value = jobplan
    sheet['C' + str(5+i)].value = description
    sheet['E' + str(5+i)].value = plannedhrs
    sheet['D' + str(5+i)].value = tstart
    sheet['F' + str(5+i)].value = asset

    if sheet['E' + str(5+i)].value == 0:
        sheet['E' + str(5+i)].value = 0.5 #If planned hours = 0 in excel, change to 0.5

        
    nextElem = browser.find_element(By.ID, "toolactions_NEXT-tbb_image") #the next work order button
    type(nextElem)
    time.sleep(30)
    try:
        nextElem.click() #click next work order button
    except:###NEW
        nextElem.click() #if that didnt work try again
    wotabElem = browser.find_element(By.ID, "mbf28cd64-tab_anchor") #work order tab button
    type(wotabElem)
    time.sleep(15)
    wotabElem.click() #click the work order tab
    time.sleep(5)
    print(WO + ' complete \n') #progress notification
    plannedhrs2 = 0
    plannedhrs3 = 0
    i=i+1

wb.save() #after the while loop ends, save, close excel, close browser.
wb.app.quit()
browser.quit()
