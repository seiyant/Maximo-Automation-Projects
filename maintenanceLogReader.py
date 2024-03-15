# maintenanceLogReader.py
# 
# This script intends to check if a daily maintenance log status matches the Maximo status
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
    # Change date, usually Wed, Thu, Sat
    file_path = 'P:\All\Maintenance\Daily Maintenance Logs\Maintenance Log Reports\Standard Test Log.xlsm'
    sheet = 'Sheet1'

    # Location of the daily maintenance log checker
    excel_file_path = 'P:\All\Maintenance\Daily Maintenance Logs\Maintenance Log Reports\Maintenance Daily Log Checker.xlsx'
    # Make a new sheet each month
    excel_file_sheet = 'March 24' # change month
    extract_and_write_excel(file_path, sheet, excel_file_path, excel_file_sheet)
    browser = webdriver.Edge()
    
    print("Web browser initiated...\n")
    extract_maximo_status(browser, excel_file_path, excel_file_sheet)
    browser.quit()
    print("Process complete...\n")

def extract_and_write_excel(file_path, sheet, excel_file_path, excel_file_sheet):
    wb = xw.Book(file_path)
    sheet = wb.sheets[sheet]

    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    print('Last row index: ', last_row)

    date = sheet.range(f'B{3}').value
    print('Date: ', date)
    
    for rows in range(2, last_row+1):
        if (sheet.range(f'A{rows}').value == 'WORK DONE') or (sheet.range(f'A{rows}').value == 'WORK STATUS'):
            work_done_index = rows
            print("Index:", work_done_index)
            break

    work_details = []
    for rows in range(work_done_index+2, last_row+1):
        name = sheet.range(f'A{rows}').value
        work_order_id = sheet.range(f'B{rows}').value
        description = sheet.range(f'C{rows}').value
        status = sheet.range(f'D{rows}').value
        work_details.append((name, work_order_id, description, status))
        print(name, work_order_id, description, status)

    wb2 = xw.Book(excel_file_path)
    sheet2 = wb2.sheets[excel_file_sheet] 

    last_row2 = sheet2.range('A' + str(sheet2.cells.last_cell.row)).end('up').row + 1
    print('Last row index: ', last_row2)

    for detail in work_details:
        name, work_order_id, description, status = detail
        print(f'{name} worked on {description} and it is {status}')

        # Writing data to Excel
        sheet2.range(f"A{last_row2}").value = date
        sheet2.range(f"B{last_row2}").value = name
        sheet2.range(f"C{last_row2}").value = work_order_id
        sheet2.range(f"D{last_row2}").value = description
        if status == "IN PROGRESS":
            status = "IP"
        sheet2.range(f"E{last_row2}").value = status

        # Increment row
        last_row2 += 1 
    
    # Save workbook after written data
    wb.save()

# Define function to fetch Maximo status using Selenium
def extract_maximo_status(browser, excel_file_path, excel_file_sheet):
    print("Extracting Maximo status...\n")
    # Connect to the workbook
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets[excel_file_sheet]

    # Find the last used row in the Excel sheet
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    
    actions = ActionChains(browser) 
    wait = WebDriverWait(browser, 10)
    browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
    browser.maximize_window()

    # Extract login information from text file
    credentials = {}
    with open('config.txt', 'r') as file:
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
    time.sleep(5)
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[7]/iframe")))
    browser.switch_to.frame(iframe)
    wotrackElem = wait.until(EC.element_to_be_clickable((By.ID, "FavoriteApp_WOTRACK")))
    actions.move_to_element_with_offset(wotrackElem, 5, 5).click().perform()
        
    # Ensure History? section is NULL
    history = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:20]_txt-tb")))
    #history.click()
    history.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    history.send_keys(Keys.ENTER)
    time.sleep(3)
    print('History updated...')
    
    # Row 1 contains headers
    for row_num in range(2, last_row + 1):
        if (sheet.range(f'F{row_num}').value != 'CLOSE'):
            # Work order is in column C
            work_order_id = sheet.range(f'C{row_num}').value

            if work_order_id:
                # Search for work order using Selenium
                searchWO_number = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:1]_txt-tb")))
                searchWO_number.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                searchWO_number.send_keys(work_order_id)
                searchWO_number.send_keys(Keys.ENTER)
                
                try:
                    time.sleep(3)
                    status = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="m6a7dfd2f_tdrow_[C:13]-c[R:0]"]/span')))
                    maximo_status = status.get_attribute('title') 
                    print(work_order_id, maximo_status)
                    
                    # Write Maximo status to Excel
                    sheet.range(f'F{row_num}').value = maximo_status
                
                except:
                    # If work order is not found in Maximo, status is "DNE"
                    sheet.range(f'F{row_num}').value = "DNE"
                    print(work_order_id, 'DNE')

            else:
                # If work order ID is missing, status is "Not Sure"
                sheet.range(f'F{row_num}').value = "NOT SURE"
                print(work_order_id, 'NOT SURE')
        
    # Save the Excel file
    print('Work orders complete...')
    wb.save()

main()
