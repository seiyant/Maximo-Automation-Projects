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

    excel_path = 'P:\All\SERVER REPORTS\PM Hours Correction.xlsx'
    excel_page = '2022'
    excel_page2 = 'Statistics'
    
    excel_saver(browser, index, total, excel_path, excel_page, excel_page2)
    
    browser.quit()

    #excel_analysis(excel_path, excel_page)

    print(f'That was {start_date} through {end_date}\n')
    print('We out')

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
    print(f'Your given time frame is {dtime.days} days')

    return skey, fkey

def maximo_navigation(browser, start_date, end_date):
    wait = WebDriverWait(browser, 20)
    #browser.get('https://test.manage.test.iko.max-it-eam.com/maximo')
    browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
    browser.maximize_window()

    # Extract login information from text file
    credentials = {}
    #with open('configtest.txt', 'r') as file:
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
    garbage_value.click()

    # Set Job Plan to GE*
    jobplange = wait.until(EC.element_to_be_clickable((By.ID, 'm6cc85995-tb')))
    jobplange.click()
    jobplange.send_keys('GE*')
    garbage_value.click()
    
    # Set Sched Start
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
    time.sleep(5)
    wostring = wait.until(EC.element_to_be_clickable((By.ID, 'm6a7dfd2f-lb3')))
    match = re.search(r'\((\d+) - (\d+) of (\d+)\)', wostring.text)

    if match:
        i = int(match.group(1))
        maxi = int(match.group(3))
    else:
        i = 0
        maxi = 0
        print('No results in this time frame. If issue persists change the sleep length')

    print(f'Total Work Orders: {maxi}')
    return i, maxi

def excel_saver(browser, index, total, path, page, page2):
    wait = WebDriverWait(browser, 20)

    wb = xw.Book(path)
    sh1 = wb.sheets[page]
    sh2 = wb.sheets[page2]

    # Change index to leave off of where stop occurred
    bypass = -113; #index = bypass 
    
    # Cycle through each line and page (index begins at 1)
    while index <= total:
        print(f'Parsing Work Order {index} of {total}')

        # Navigate to Work Order
        if index == bypass: # Enter WO and bypass all viewed WOs
            wo_html = f'm6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]'
            findwo = wait.until(EC.element_to_be_clickable((By.ID, wo_html)))
            findwo.click()
            time.sleep(2)
            p = 1
            while p != index:
                print(f'Moving to WO #{p}')
                try:
                    findwo = wait.until(EC.element_to_be_clickable((By.ID, 'toolactions_NEXT-tbb_anchor')))
                    findwo.click()
                    time.sleep(2)
                    nav2wo = wait.until(EC.element_to_be_clickable((By.ID,'mbf28cd64-tab_anchor')))
                    nav2wo.click()
                    time.sleep(1)

                except:
                    try:
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
                p += 1

                
        elif index == 1: # Enter first WO
            wo_html = f'm6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:{index - 1}]'
            findwo = wait.until(EC.element_to_be_clickable((By.ID, wo_html)))
            findwo.click()
        
        else:
            try:
                findwo = wait.until(EC.element_to_be_clickable((By.ID, 'toolactions_NEXT-tbb_anchor')))
                findwo.click()
                time.sleep(3)
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

        # Navigate in Work Order
        time.sleep(2)
        work_order = wait.until(EC.element_to_be_clickable((By.ID, 'm52945e17-tb'))).get_attribute('value')
        print(f'Work Order: {work_order}')
        description = wait.until(EC.element_to_be_clickable((By.ID, 'md42b94ac-tb'))).get_attribute('value')
        print(f'Description: {description}')
        work_type = wait.until(EC.element_to_be_clickable((By.ID, 'me2096203-tb'))).get_attribute('value')
        print(f'Work Type: {work_type}')
        job_plan = wait.until(EC.element_to_be_clickable((By.ID, 'mfe7bb84-tb'))).get_attribute('value')
        print(f'Job Plan: {job_plan}')
        duration_hr, duration_min = map(int, wait.until(EC.element_to_be_clickable((By.ID, 'm8c7fa385-tb'))).get_attribute('value').split(':'))
        plan_hours = duration_hr + (duration_min / 60)
        print(f'Planned Hours: {plan_hours}')
        if plan_hours == 0.0:
            print('0 Planned Hours Error: Auto-rounding to 15 minutes')
            plan_hours = 0.25
            print(f'Planned Hours: {plan_hours}')

        # Navigate to Plans
        plans = wait.until(EC.element_to_be_clickable((By.ID, 'm356798d1-tab_anchor')))
        plans.click()
        time.sleep(3)
        plan_qty = wait.until(EC.element_to_be_clickable((By.ID, 'm5e4b62f0_tdrow_[C:6]_txt-tb[R:0]'))).get_attribute('value')
        print(f'Planned Quantity: {plan_qty}')
        
        # Navigate to Actuals
        actuals = wait.until(EC.element_to_be_clickable((By.ID, 'm272f5640-tab_anchor')))
        actuals.click()

        # There may be multiple laborers
        time.sleep(3)
        
        laborstring = wait.until(EC.element_to_be_clickable((By.ID, 'm4dfd8aef-lb3')))
        match = re.search(r'\((\d+) - (\d+) of (\d+)\)', laborstring.text)
        labor_max = int(match.group(3))
        
        # Labor page adjusting
        while int(match.group(1)) != 1:
            if int(match.group(1)) == 0:
                # Since labor_max is 0 the breaking this loop will restart the index while loop
                print(f'No laborers error, skipping work order...')
                break
            else:
                laborstring = wait.until(EC.element_to_be_clickable((By.ID, 'm4dfd8aef-lb3')))
                match = re.search(r'\((\d+) - (\d+) of (\d+)\)', laborstring.text)
                print(f'Adjusting page...')
                adjust_laborpage = wait.until(EC.element_to_be_clickable((By.ID, 'm4dfd8aef-ti6')))
                adjust_laborpage.click()
                time.sleep(3)
        
        # Labor saving
        j = 0; adjust_j = 0; total_real_hours = 0
        labor_hours = {}

        while j < labor_max:
            time.sleep(2)
            print(f'Index {j + 1} of {labor_max}')

            labor_xpath = f'm4dfd8aef_tdrow_[C:3]_txt-tb[R:{j}]'
            hours_xpath = f'm4dfd8aef_tdrow_[C:9]_txt-tb[R:{j}]'

            laborer = wait.until(EC.element_to_be_clickable((By.ID, labor_xpath))).get_attribute('value')
            hours_hr, hours_min = map(int, wait.until(EC.element_to_be_clickable((By.ID, hours_xpath))).get_attribute('value').split(':'))
            real_hours = hours_hr + (hours_min / 60)

            if real_hours != 0:
                print(f'{laborer} worked for {real_hours} hours')
                total_real_hours += real_hours 
                if laborer in labor_hours:
                    labor_hours[laborer] += real_hours
                else: # Add new laborer
                    labor_hours[laborer] = real_hours 
            else: 
                print('0 Hours Worked Error: Entry will be skipped')
                adjust_j += 1
            
            if j % 6 == 5:
                flip_laborpage = wait.until(EC.element_to_be_clickable((By.ID, 'm4dfd8aef-ti7')))
                flip_laborpage.click()
                time.sleep(3)
            
            j += 1
        
        # Loop for data entry if labor is not 0
        if (labor_max - adjust_j) > 0:
            # Excel Magic
            last_row1 = sh1.range('I' + str(sh1.cells.last_cell.row)).end('up').row + 1
            last_row2 = sh2.range('A' + str(sh2.cells.last_cell.row)).end('up').row
            print(f'Next row: {last_row1}')

            sh1.range(f'A{last_row1}').value = job_plan
            sh1.range(f'B{last_row1}').value = work_type
            sh1.range(f'C{last_row1}').value = description
            sh1.range(f'D{last_row1}').value = work_order
            sh1.range(f'E{last_row1}').value = plan_hours
            sh1.range(f'F{last_row1}').value = total_real_hours

            # Statistics section entry
            labor_column = 7 # 7 is column G
            for laborer, labor_hours in labor_hours.items():
                labor_cell = f'{chr(64 + labor_column)}{last_row1}' # A is 64 in ASCII
                sh1.range(labor_cell).value = laborer
                labor_column += 1
                
                # Statistics Section
                laborer_found = False
                search_range = sh2.range(f'A1:A' + str(last_row2))
                
                for cell in search_range:
                    if cell.value == laborer:
                        laborer_found = True
                        current_row = cell.row + 2  # Start looking after the headers
                        
                        # Insert the new row before the next laborer's name (if exists) or in the next empty row
                        sh2.range(f'A{current_row}:F{current_row}').api.Insert(Shift=xw.constants.InsertShiftDirection.xlShiftDown)
                        sh2.range(f'A{current_row}:F{current_row}').value = [job_plan, work_type, description, work_order, (int(plan_hours) / int(plan_qty)), labor_hours]
                        break
                if not laborer_found:
                    # Laborer not found, add at the end
                    last_row2 = sh2.range('A' + str(sh2.cells.last_cell.row)).end('up').row + 2  # Update the last_row2 to reflect the new last row
                    sh2.range(f'A{last_row2}').value = laborer
                    sh2.range(f'A{last_row2 + 1}').value = ["Job Plan", "Work Type", "Description", "Work Order", "Planned Hours", "Actual Hours"]
                    sh2.range(f'A{last_row2 + 2}').value = [job_plan, work_type, description, work_order, (int(plan_hours) / int(plan_qty)), labor_hours]
#add in % difference and prob dist function for statistics

        print('\n')
        index += 1
        time.sleep(2)
main()
