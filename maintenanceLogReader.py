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
import xlwings as xw

# Assumes maintenance log is in the same folder
print("Make sure daily maintenance log is in the same project folder...\n")
print("Excel file: Maintenance Daily Log Checker\n")
print("Sheet: List of Records\n")
word_name = input("Copy and paste the Word document name\n")
word_file_path = word_name".doc"
formatted_date, work_details = extract_data_from_doc(word_file_path)
write_to_excel(formatted_date, work_details)
browser = webdriver.Edge()
extract_maximo_status(browser)
browser.quit()
print("Process complete...\n")

# Extract the work details from the Word document
def extract_data_from_doc(file_path):
    # Load docx file
    docx = Document(file_path)

    # Extract the date and format it
    date_cell = docx.tables[0].cell(1, 1).text
    date_object = datetime.datetime.strptime(date_cell, '%b %d, %Y')
    formatted_date = date_object.strftime('%m/%d/%Y')  

    # Initialize an empty list to store the work details
    work_details = []

    # Extract crew attendance details
    crew_names = {}
    nickname_initial_mapping = {
        'William': 'B',
        'Will': 'B',
        'Robert': 'B'
    }
    for row in docx.tables[0].rows:
        if row.cells[0].text == "Crew":
            for crew_row in docx.tables[0].rows:
                name = crew_row.cells[0].text
                position = crew_row.cells[1].text
                attendance = crew_row.cells[3].text
                if attendance != "A":
                    first_name, last_name = name.split()
                    initials = first_name[0] + last_name[0]
                    crew_names[initials] = name

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

            # Check if the initials match with crew names
            if job_assigned_to == "Everyone":
                name = "Everyone"
            else:
                assigned_names = []
                # Split the initials by "/" and iterate through each one
                for initial in job_assigned_to.split("/"):
                    if initial in crew_names:
                        assigned_names.append(crew_names[initial])
                    else:
                        # Check if name is one of the names with possible nickname
                        for first_name, new_initial in nickname_initial_mapping.items():
                            possible_initials = new_initial + job_assigned_to[1]
                            if possible_initials in crew_names and first_name in crew_names[possible_initials]:
                                assigned_names.append(crew_names[possible_initials])
                                break
                    # else look in "laborAssignment.xls" under "List of Records" sheet in column "H", called "Initials" and check for initials (without the nickname) matching. List the options in the form of names if multiple appear and let user choose out of options.

                name = "/".join(assigned_names)
            work_details.append((name, description, status))

    return formatted_date, work_details

# Define function to write to Excel
def write_to_excel(formatted_date, work_details):
    # Connect to the workbook
    wb = xw.Book(excel_path)
    sheet = wb.sheets['Automated']

    # Find the last used row in the Excel sheet
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

    # Start entering data from the next available row
    for detail in work_details:
        name, job_assigned_to, description, status = detail

        # Split the description to extract 9-digit ID
        work_order_id = None
        for word in description.split():
            if word.startswith("W23") and len(word) == 9:
                work_order_id = word
                break
            if work_order_id:
                description = description.replace(work_order_id, '').strip()

        # Increment row
        last_row += 1 

        # Writing data to Excel
        sheet.range(f"A{last_row}").value = formatted_date
        sheet.range(f"B{last_row}").value = name
        sheet.range(f"C{last_row}").value = work_order_id
        sheet.range(f"D{last_row}").value = description
        sheet.range(f"E{last_row}").value = status
    
    # Save workbook after written data
    wb.save()

# Define function to fetch Maximo status using Selenium
def extract_maximo_status(browser, excel_path):
    # Connect to the workbook
    wb = xw.Book(excel_path)
    sheet = wb.sheets['Automated']

    # Find the last used row in the Excel sheet
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    
    # Navigate to Maximo login page
    actions = ActionChains(browser) 
    wait = WebDriverWait(browser, 20)
    browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
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

    # Navigate to Work Order Tracking
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[7]/iframe")))
    browser.switch_to.frame(iframe)
    wotrackElem = wait.until(EC.element_to_be_clickable((By.ID, "FavoriteApp_WOTRACK")))
    actions.move_to_element_with_offset(wotrackElem, 5, 5).click().perform()
    
    # Ensure History? section is NULL
    history = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:20]_txt-tb")))
    history.click()
    history.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    
    # Row 1 contains headers
    for row_num in range(2, last_row + 1):  
        # Work order is in column C
        work_order_id = sheet.range(f'C{row_num}').value 

        if work_order_id:
            # Search for work order using Selenium
            searchWO_number = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:1]_txt-tb")))
            searchWO_number.send_keys(work_order_id)
            searchWO_number.send_keys(Keys.ENTER)
            
            try:
                status_element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="m6a7dfd2f_tdrow_[C:13]-c[R:0]"]/span')))
                maximo_status = status_element.text
                
                # Write Maximo status to Excel
                sheet.range(f'F{row_num}').value = maximo_status
            
            except:
                # If work order is not found in Maximo, status is "DNE"
                sheet.range(f'F{row_num}').value = "DNE"

        else:
            # If work order ID is missing, status is "Not Sure"
            sheet.range(f'F{row_num}').value = "Not Sure"
    
    # Save the Excel file
    wb.save()