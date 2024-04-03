# plannedHoursCorrectionEntry
# 
# This script intends to corrected hours data back into Maximo
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

def main():
    excel_path = 'P:\All\SERVER REPORTS\PM Hours Correction.xlsx'
    excel_page = '2020' # Use 'Overall'

    browser = webdriver.Edge()
    wait = WebDriverWait(browser, 20)

    browser_login(browser, wait)

    excel_reader(excel_page, excel_path, browser, wait)

    browser.quit()

def browser_login(browser, wait):
    #browser.get('https://test.manage.test.iko.max-it-eam.com/maximo')
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

    # Navigate to Job Plans (this is wrong right now)
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[7]/iframe')))
    browser.switch_to.frame(iframe)
    planningElem = wait.until(EC.element_to_be_clickable((By.ID, 'm7f8f3e49_ns_menu_PLANS_MODULE_a')))
    planningElem.click()
    jobplanElem = wait.until(EC.element_to_be_clickable((By.ID, 'm7f8f3e49_ns_menu_PLANS_MODULE_sub_changeapp_JOBPLAN_a')))
    jobplanElem.click()
    

def excel_reader(excel_page, excel_path, browser, wait):
    wb = xw.Book(excel_path)
    sh = wb.sheets[excel_page]

    # Ensure Excel is pre-sorted, JP -> WO
    last_row = sh.range('A' + str(sh.cells.last_cell.row)).end('up').row
    row_skip = 4
    
    jp_last = 0; ro_count = 0; wo_last = 0; wo_count = 0; hr_sum = 0; wk_count = 0
    for row in range(row_skip, last_row + 1):
        wo = sh.range(f'D{row}').value
        jp = sh.range(f'A{row}').value
        hr = sh.range(f'F{row}').value
        lq = sh.range(f'F{row}').value

        print(f'\nRow {row}')
        print(f'Work Order: {wo}')
        print(f'Job Plan: {jp}')
        print(f'Hours Worked: {hr}')
        print(f'Labor Quantity: {lq}')

        if ((jp == jp_last) | (row == row_skip)) & (row != last_row + 1): #includes first occurrence
            ro_count += 1

            if wo == wo_last: # nothing changes
                wo_last = wo # values are equal
                print('Same JP and same WO')

            elif row == row_skip:
                jp_last = jp
                wo_last = wo
                hr_sum += hr
                wo_count += 1
                print('First row')

            else: 
                wo_count += 1
                hr_sum += hr
                wo_last = wo # reset work order
                print('Same JP and different WO')

        else: #if last row go here
            print('Different JP\n')
            # Floor divide average number of laborers used
            wk_count = ro_count // wo_count 
            hr_avg = hr_sum / wo_count
            hr_ind = hr_avg / wk_count

            # Look up Job Plan
            print(f'Searching {jp_last} in Maximo...')
            find_jp = wait.until(EC.element_to_be_clickable((By.ID, 'm6a7dfd2f_tfrow_[C:1]_txt-tb')))
            find_jp.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
            find_jp.send_keys(jp_last, Keys.ENTER)
            time.sleep(2)

            # Enter Job Plan
            click_jp = wait.until(EC.element_to_be_clickable((By.ID, 'm6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]')))
            click_jp.click()
            time.sleep(2)

            # Set Duration
            print(f'Separate Laborers: {ro_count}')
            print(f'WO Count: {wo_count}')
            print(f'Average Laborers: {wk_count}')
            print(f'Planned hours changing to {hr_avg} hours with {wk_count} laborers, {hr_ind} hours each')
            duration = wait.until(EC.element_to_be_clickable((By.ID, 'maa8ad01-tb')))
            duration.click()
            duration.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
            duration.send_keys(hr_avg)

            # Garbage value to ensure things load
            garbage_value = wait.until(EC.element_to_be_clickable((By.ID, 'mff46efd3-tb')))
            garbage_value.click()
            time.sleep(1)

            # Set Quantity
            quantity = wait.until(EC.element_to_be_clickable((By.ID, 'ma0e8b2fb_tdrow_[C:7]_txt-tb[R:0]')))
            quantity.click()
            quantity.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
            quantity.send_keys(wk_count)
            garbage_value = wait.until(EC.element_to_be_clickable((By.ID, 'ma0e8b2fb_tdrow_[C:5]_txt-tb[R:0]')))
            garbage_value.click()
            time.sleep(1)
            
            # Set Hours
            hours = wait.until(EC.element_to_be_clickable((By.ID, 'ma0e8b2fb_tdrow_[C:8]_txt-tb[R:0]')))
            hours.click()
            hours.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
            hours.send_keys(hr_ind)

            # Save Work
            save = wait.until(EC.element_to_be_clickable((By.ID, 'toolactions_SAVE-tbb_anchor')))
            save.click()
            time.sleep(2)
            
            # Exit 
            menu = wait.until(EC.element_to_be_clickable((By.ID, 'mab323381_nc_list_button')))
            menu.click()

            # Resets for new job plan
            jp_last = jp
            wo_count = 1
            hr_sum = hr
            ro_count = 1

main()
