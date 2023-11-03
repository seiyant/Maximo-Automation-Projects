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
import re
import datetime
import xlwings as xw

# Extract the work details from the Word document
def extract_data_from_doc(file_path):
    # Load docx file
    doc = Document(file_path)
    print('Loaded Word document...\n')
    
    # Initialize an empty list to store the work details
    work_details = []

    # Find cell locations
    row_index = 0
    for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if (''.join((cell.text).split()) == 'DATE') and (cell.text != last_cell_text):
                        date_row = row_index # Date is in the same row
                        print('Date:', date_row)
                    if (''.join((cell.text).split()) == 'ERTORLEVEL3FIRSTAID') and (cell.text != last_cell_text):
                        crew_row = row_index + 1
                        print('Crew:', crew_row)
                    if (''.join((cell.text).split()) == 'STATUS') and (cell.text != last_cell_text):
                        jobs_row = row_index + 1
                        print('Jobs:', jobs_row)
                    last_cell_text = cell.text
                row_index += 1
    print('Date row index:', date_row)
    print('Crew row index:', crew_row)
    print('Jobs row index:', jobs_row, '\n')

    # Extract the date and format it
    for cell in doc.tables[0].rows[date_row].cells:
        if (''.join((cell.text).split()) != 'DATE') and (cell.text != last_cell_text):
            date = cell.text
            break
        last_cell_text = cell.text
    print('DATE: ', date, '\n')
    
    # Extract crew attendance details
    crew_names = {}
    crew_positions = {}
    cell_tracker = 0 # Keep track of index
    end_of_loop = False
    
    for cell in doc.tables[0].rows[crew_row - 1].cells:
        if ''.join((cell.text).split()) == 'CREW':
            crew_count = cell_tracker
        elif ''.join((cell.text).split()) == 'POSITION':
            position_count = cell_tracker
        elif ''.join((cell.text).split()) == 'PRESENT(P)/ABSENT(A)':
            attendance_count = cell_tracker
        elif ''.join((cell.text).split()) == 'ERTORLEVEL3FIRSTAID':
            certification_count = cell_tracker # Not really important
        cell_tracker += 1
    print('Crew cell size:', crew_count)
    print('Position cell size:', position_count)
    print('Attendance cell size:', attendance_count)
    print('Certification cell size:', certification_count, '\n')

    for row in doc.tables[0].rows[crew_row:]:
        if row.cells[attendance_count].text == 'P':
            first_name, last_name = row.cells[crew_count].text.split()
            position = row.cells[position_count].text
            
            real_initials = first_name[0] + last_name[0]
            crew_names[real_initials] = row.cells[crew_count].text
            crew_positions[real_initials] = position
            
            print(first_name, last_name, position)
            
        for cell in row.cells:
            if ''.join((cell.text).split()) == 'JOBASSIGNEDTO':
                end_of_loop = True
                break  
        if end_of_loop == True:
            break
        
    print('\n')

    # Consider common letter swapping nicknames
    nickname_initials = {
        'William': 'B',
        'Will': 'B',
        'Robert': 'B'
    }

    # Extracts work order details
    cell_tracker = 0 # Keep track of index
    empty_cell_1 = False
    empty_cell_2 = False
    
    for row in doc.tables[0].rows[jobs_row:]:
        for cell in row.cells:
            if cell.text != last_cell_text:
                if cell_tracker == 0:
                    job_assigned_to = cell.text
                elif cell_tracker == 1:
                    description = cell.text
                    if ''.join((cell.text).split()) == '':
                        empty_cell_1 = True 
                elif cell_tracker == 2:
                    status = cell.text
                    if ''.join((cell.text).split()) == '':
                        empty_cell_2 = True
                    if job_assigned_to.lower() in ['everyone', 'all', '']: # Special case
                        name = 'Everyone' 
                    else:
                        assigned_names = []
                        for initials in job_assigned_to.split('/'):
                            if initials in crew_names: # Regular search
                                assigned_names.append(crew_names[initials])
                            else:
                                matched_initials = False
                                for first_nickname, nickname_initial in nickname_initials.items():
                                    possible_initials = nickname_initial + initials[1]
                                    if possible_initials in crew_names: # Search with letter swap
                                        assigned_names.append(crew_names[possible_initials])
                                        matched_initials = True
                                        break
                                if not matched_initials:
                                    excel_initials = search_in_excel(initials) # Search in database
                                    
                        name = ', '.join(assigned_names)
                        
                    work_details.append((name, description, status))
                    print('Crew: ', name, '\nDescription: ', description, '\nStatus: ', status)
                    cell_tracker = -1 # Reset index
                    
                last_cell_text = cell.text
                cell_tracker += 1
                
            if (empty_cell_1 == True) and (empty_cell_2 == True):
                break
            
    return (date, work_details)

# Function to extract data from laborAssignments as a fail case
def search_in_excel(initial):
    print(f"Searching in Excel for initial: {initial}...\n")
    # Load Excel file
    workbook = xw.Book('laborAssignment.xlsx')
    sheet = workbook.sheets['List of Records']

    # Find the last row with data in column H
    last_row = sheet.range('H' + str(sheet.cells.last_cell.row)).end('up').row

    # Execute all initials from column H, H1 is the subtitle
    all_initials = sheet.range('H2:H' + str(last_row)).value

    # Check for matches based on provided initials and potential nickname initials
    potential_initials = [initial]
    if initial[0] == 'B':
        potential_initials.extend('W' + initial[1], 'R' + initial[1])

    # If initial is found, get corresponding name and position
    matching_rows = [index for index, value in enumerate(all_initials) if value in potential_initials]

    results = []
    for row in matching_rows:
        name = sheet.range('B' + str(row + 2)).value
    
    workbook.close()

    '''if excel_results:
                                # If multiple matches, prompt user to choose
                                if len(excel_results) > 1:
                                    print(f"Multiple matches found for initials {initial}:")
                                    for index, (name_option, _) in enumerate(excel_results):
                                        print(f"{index + 1}. {name_option}")
                                    choice = int(input("Choose the correct match (enter the number): ")) - 1
                                    assigned_names.append(excel_results[choice][0])
                                else:
                                    assigned_names.append(excel_results[0][0])
                            else:
                                assigned_names.append("Name DNE")'''
                                
    return name

# Define function to write to Excel
def write_to_excel(date_cell, work_details):
    print("Writing extracted data to Excel...\n")
    # Connect to the workbook
    wb = xw.Book('Maintenance Daily Log Checker.xlsx')
    sheet = wb.sheets['September 23'] # Change every Month

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
        sheet.range(f"A{last_row}").value = date_cell
        sheet.range(f"B{last_row}").value = name
        sheet.range(f"C{last_row}").value = work_order_id
        sheet.range(f"D{last_row}").value = description
        sheet.range(f"E{last_row}").value = status
    
    # Save workbook after written data
    wb.save()

# Define function to fetch Maximo status using Selenium
def extract_maximo_status(browser, excel_path):
    print("Extracting Maximo status...\n")
    # Connect to the workbook
    wb = xw.Book(excel_path)
    sheet = wb.sheets['September 23']

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

# Assumes maintenance log is in the same folder
print("Initializing...\n")
print("Make sure daily maintenance log is in the same project folder...\n")
print("Excel file: Maintenance Daily Log Checker.xlsx...\n")
print("Sheet in Excel file: List of Records...\n")
word_name = input("Enter the date (MTH D, YEAR) of Word document:\n")
word_file_path = word_name + " Maintenance Daily Log.docx"
date_cell, work_details = extract_data_from_doc(word_file_path)
print(f"Total work details extracted: {len(work_details)}\n")
write_to_excel(date_cell, work_details) # Load excel from here not inside
browser = webdriver.Edge()
print("Web browser initiated...\n")
extract_maximo_status(browser, 'Maintenance Daily Log Checker.xlsx')
browser.quit()
print("Process complete...\n")
