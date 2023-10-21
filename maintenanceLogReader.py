# InstantMaximo.py
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
import datetime
import xlwings as Excel

# Ensure browser version and web driver version match
browser = webdriver.Edge()
actions = ActionChains(browser) 

# Define wait
wait = WebDriverWait(browser, 20)

# Navigate to Maximo login page
browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   

# Maximize window
browser.maximize_window()

# Enter login information
UserElem = wait.until(EC.element_to_be_clickable((By.ID, "username")))
UserElem.send_keys('NOZASEIY')

Cont1Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
Cont1Elem.click()

passElem = wait.until(EC.element_to_be_clickable((By.ID, "password")))
passElem.send_keys('Roofing1SN!') 

Cont2Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
Cont2Elem.click()

# Navigate to Quick Reporting
iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[7]/iframe")))
browser.switch_to.frame(iframe)

# Extract the work details from the Word document
def extract_data_from_docx(file_path):
    # Load docx file
    docx = Document(file_path)

    # Extract the date and format it
    date_cell = docx.tables[0].cell(1, 1).text
    date_object = datetime.datetime.strptime(date_cell, '%b %d, %Y')
    formatted_date = date_object.strftime('%m/%d/%Y')  

    # Initialize an empty list to store the work details
    work_details = []

    # Search for the row that contains "JOBS ASSIGNED TO" in column 0
    start_row_index = None
    for index, row in enumerate(docx.tables[0].rows):
        if row.cells[0].text.lower() == "JOB ASSIGNED TO":
            start_row_index = index
            break
    
    # If "JOB ASSIGNED TO" is found, extract subsequent rows
    if start_row_index is not None:
        for row in docx.tables[0].rows[start_row_index + 1:]:
            job_assigned_to = row.cells[0].text
            description = row.cells[1].text
            status = row.cells[2].text
            work_details.append((job_assigned_to, description, status))

    return formatted_date, work_details

# Cross-reference with Excel
def get_full_name(excel_path):
    # Open Excel workbook and select sheet
    wb = Excel.Book(excel_path)
    sheet = wb.sheets[0]

    # Get full name (column 1) and initials (column 4)
    full_name = sheet.range('B1:B' + str(sheet.cells.last_cell.row)).value
    initials = sheet.range('E1:E' + str(sheet.cells.last_cell.row)).value
    
    # Create a disctionary to map initials to full names
    initials_dictionary = {initial: name for initial, name in zip(initials, full_names)}

    return initials_dictionary

def name_to_code(work_details, sheet, crew):
    # Dictionary 

# Define function to fetch Maximo status using Selenium
# Define function to write to Excel
# Main function or script execution
# Close Selenium driver and save Excel file